#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; Other includes for Writer

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Deleting L.O. Writer Forms and Form Controls.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_FormAdd
; _LOWriter_FormConCheckBoxData
; _LOWriter_FormConCheckBoxGeneral
; _LOWriter_FormConCheckBoxState
; _LOWriter_FormConComboBoxData
; _LOWriter_FormConComboBoxGeneral
; _LOWriter_FormConComboBoxValue
; _LOWriter_FormConCurrencyFieldData
; _LOWriter_FormConCurrencyFieldGeneral
; _LOWriter_FormConCurrencyFieldValue
; _LOWriter_FormConDateFieldData
; _LOWriter_FormConDateFieldGeneral
; _LOWriter_FormConDateFieldValue
; _LOWriter_FormConDelete
; _LOWriter_FormConFileSelFieldGeneral
; _LOWriter_FormConFileSelFieldValue
; _LOWriter_FormConFormattedFieldData
; _LOWriter_FormConFormattedFieldGeneral
; _LOWriter_FormConFormattedFieldValue
; _LOWriter_FormConGetParent
; _LOWriter_FormConGroupBoxGeneral
; _LOWriter_FormConImageButtonGeneral
; _LOWriter_FormConImageControlData
; _LOWriter_FormConImageControlGeneral
; _LOWriter_FormConInsert
; _LOWriter_FormConLabelGeneral
; _LOWriter_FormConListBoxData
; _LOWriter_FormConListBoxGeneral
; _LOWriter_FormConListBoxGetCount
; _LOWriter_FormConListBoxSelection
; _LOWriter_FormConNavBarGeneral
; _LOWriter_FormConNumericFieldData
; _LOWriter_FormConNumericFieldGeneral
; _LOWriter_FormConNumericFieldValue
; _LOWriter_FormConOptionButtonData
; _LOWriter_FormConOptionButtonGeneral
; _LOWriter_FormConOptionButtonState
; _LOWriter_FormConPatternFieldData
; _LOWriter_FormConPatternFieldGeneral
; _LOWriter_FormConPatternFieldValue
; _LOWriter_FormConPosition
; _LOWriter_FormConPushButtonGeneral
; _LOWriter_FormConPushButtonState
; _LOWriter_FormConsGetList
; _LOWriter_FormConSize
; _LOWriter_FormConTableConCheckBoxData
; _LOWriter_FormConTableConCheckBoxGeneral
; _LOWriter_FormConTableConColumnAdd
; _LOWriter_FormConTableConColumnDelete
; _LOWriter_FormConTableConColumnsGetList
; _LOWriter_FormConTableConComboBoxData
; _LOWriter_FormConTableConComboBoxGeneral
; _LOWriter_FormConTableConCurrencyFieldData
; _LOWriter_FormConTableConCurrencyFieldGeneral
; _LOWriter_FormConTableConDateFieldData
; _LOWriter_FormConTableConDateFieldGeneral
; _LOWriter_FormConTableConFormattedFieldData
; _LOWriter_FormConTableConFormattedFieldGeneral
; _LOWriter_FormConTableConGeneral
; _LOWriter_FormConTableConListBoxData
; _LOWriter_FormConTableConListBoxGeneral
; _LOWriter_FormConTableConNumericFieldData
; _LOWriter_FormConTableConNumericFieldGeneral
; _LOWriter_FormConTableConPatternFieldData
; _LOWriter_FormConTableConPatternFieldGeneral
; _LOWriter_FormConTableConTextBoxData
; _LOWriter_FormConTableConTextBoxGeneral
; _LOWriter_FormConTableConTimeFieldData
; _LOWriter_FormConTableConTimeFieldGeneral
; _LOWriter_FormConTextBoxCreateTextCursor
; _LOWriter_FormConTextBoxData
; _LOWriter_FormConTextBoxGeneral
; _LOWriter_FormConTimeFieldData
; _LOWriter_FormConTimeFieldGeneral
; _LOWriter_FormConTimeFieldValue
; _LOWriter_FormDelete
; _LOWriter_FormGetObjByIndex
; _LOWriter_FormParent
; _LOWriter_FormPropertiesData
; _LOWriter_FormPropertiesGeneral
; _LOWriter_FormsGetCount
; _LOWriter_FormsGetList
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormAdd
; Description ...: Add a form to a Document or create a sub-form.
; Syntax ........: _LOWriter_FormAdd(ByRef $oObj, $sName)
; Parameters ....: $oObj                - [in/out] an object. Either a Document Object or a Form object. See Remarks. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function, or a Form Object returned from a previous _LOWriter_FormsGetList, or _LOWriter_FormAdd function.
;                  $sName               - a string value. The name of the new Form. Form names do not need to be unique, but it will help for clarity.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Called Object in $oObj, not a Document Object, and not a Form Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.form.component.Form" object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve parent Document Object.
;                  @Error 3 @Extended 2 Return 0 = Parent Document is Read Only.
;                  @Error 3 @Extended 3 Return 0 = Failed to insert new Form.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve new Form Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to name Form.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning newly created Form's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $oObj is called with a Document object, the new form will be a top level form. If a Form object is called, a sub-form will be created.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormAdd(ByRef $oObj, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oForm, $oDoc, $oInsertObj
	Local $sTempName = "AutoIt_FORM_"
	Local $iCount = 1

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oObj.supportsService("com.sun.star.form.component.Form") Then
		$oDoc = $oObj ; Identify the parent document.

		Do
			$oDoc = $oDoc.getParent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		Until $oDoc.supportsService("com.sun.star.text.TextDocument")

		$oInsertObj = $oObj

	ElseIf $oObj.supportsService("com.sun.star.text.TextDocument") Then
		$oDoc = $oObj
		$oInsertObj = $oObj.DrawPage.Forms()

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; wrong type of input item.
	EndIf

	If $oDoc.IsReadOnly() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oForm = $oDoc.createInstance("com.sun.star.form.component.Form")
	If Not IsObj($oForm) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oInsertObj.hasByName($sTempName & $iCount)
		$iCount += 1
	WEnd

	$oInsertObj.insertByName($sTempName, $oForm)
	If Not $oInsertObj.hasByName($sTempName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oForm = $oInsertObj.getByName($sTempName)
	If Not IsObj($oForm) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$oForm.Name = $sName
	If ($oForm.Name() <> $sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oForm)
EndFunc   ;==>_LOWriter_FormAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConCheckBoxData
; Description ...: Set or Retrieve Check Box Data Properties.
; Syntax ........: _LOWriter_FormConCheckBoxData(ByRef $oCheckBox[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oCheckBox           - [in/out] an object. A Check Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCheckBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oCheckBox not a Check Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
;                  Reference Values are not included here as they are applicable to Calc only, as far as I can ascertain.
; Related .......: _LOWriter_FormConCheckBoxState, _LOWriter_FormConCheckBoxGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConCheckBoxData(ByRef $oCheckBox, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oCheckBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oCheckBox) <> $LOW_FORM_CON_TYPE_CHECK_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oCheckBox.Control.DataField(), $oCheckBox.Control.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCheckBox.Control.DataField = $sDataField
		$iError = ($oCheckBox.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCheckBox.Control.InputRequired = $bInputRequired
		$iError = ($oCheckBox.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConCheckBoxData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConCheckBoxGeneral
; Description ...: Set or Retrieve general Checkbox control properties.
; Syntax ........: _LOWriter_FormConCheckBoxGeneral(ByRef $oCheckBox[, $sName = Null[, $sLabel = Null[, $oLabelField = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bVisible = Null[, $bPrintable = Null[, $bTabStop = Null[, $iTabOrder = Null[, $iDefaultState = Null[, $mFont = Null[, $iStyle = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $bWordBreak = Null[, $sGraphics = Null[, $iGraphicAlign = Null[, $bTriState = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oCheckBox           - [in/out] an object. A Checkbox Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $iDefaultState       - [optional] an integer value (0-2). Default is Null. The Default state of the Checkbox, $LOW_FORM_CON_CHKBX_STATE_NOT_DEFINED is only available if $bTriState is True. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iStyle              - [optional] an integer value (1-2). Default is Null. The display style of the checkbox. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bWordBreak          - [optional] a boolean value. Default is Null. If True, line breaks are allowed to be used.
;                  $sGraphics           - [optional] a string value. Default is Null. The path to an Image file.
;                  $iGraphicAlign       - [optional] an integer value (0-12). Default is Null. The Alignment of the Image. See Constants $LOW_FORM_CON_IMG_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTriState           - [optional] a boolean value. Default is Null. If True, the checkbox will have a third checked state.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCheckBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oCheckBox not a Check Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 6 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 7 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 13 Return 0 = $iDefaultState not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 14 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 15 Return 0 = $iStyle not an Integer, less than 1 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 16 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 17 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 18 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 19 Return 0 = $bWordBreak not a Boolean.
;                  @Error 1 @Extended 20 Return 0 = $sGraphics not a String.
;                  @Error 1 @Extended 21 Return 0 = $iGraphicAlign not an Integer, less than 0 or greater than 12. See Constants $LOW_FORM_CON_IMG_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 22 Return 0 = $bTriState not a Boolean.
;                  @Error 1 @Extended 23 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 24 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 25 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $oLabelField
;                  |                               8 = Error setting $iTxtDir
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bVisible
;                  |                               64 = Error setting $bPrintable
;                  |                               128 = Error setting $bTabStop
;                  |                               256 = Error setting $iTabOrder
;                  |                               512 = Error setting $iDefaultState
;                  |                               1024 = Error setting $mFont
;                  |                               2048 = Error setting $iStyle
;                  |                               4096 = Error setting $iAlign
;                  |                               8192 = Error setting $iVertAlign
;                  |                               16384 = Error setting $iBackColor
;                  |                               32768 = Error setting $bWordBreak
;                  |                               65536 = Error setting $sGraphics
;                  |                               131072 = Error setting $iGraphicAlign
;                  |                               262144 = Error setting $bTriState
;                  |                               524288 = Error setting $sAddInfo
;                  |                               1048576 = Error setting $sHelpText
;                  |                               2097152 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 22 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $sGraphics is called with an invalid Graphic URL, graphic is set to Null. The Return for $sGraphics is an Image Object.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $iDefaultState, $mFont, $sAddInfo.
; Related .......: _LOWriter_FormConCheckBoxState, _LOWriter_FormConCheckBoxData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConCheckBoxGeneral(ByRef $oCheckBox, $sName = Null, $sLabel = Null, $oLabelField = Null, $iTxtDir = Null, $bEnabled = Null, $bVisible = Null, $bPrintable = Null, $bTabStop = Null, $iTabOrder = Null, $iDefaultState = Null, $mFont = Null, $iStyle = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $bWordBreak = Null, $sGraphics = Null, $iGraphicAlign = Null, $bTriState = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[22]

	If Not IsObj($oCheckBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ( __LOWriter_FormConIdentify($oCheckBox) <> $LOW_FORM_CON_TYPE_CHECK_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $oLabelField, $iTxtDir, $bEnabled, $bVisible, $bPrintable, $bTabStop, $iTabOrder, $iDefaultState, $mFont, $iStyle, $iAlign, $iVertAlign, $iBackColor, $bWordBreak, $sGraphics, $iGraphicAlign, $bTriState, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oCheckBox.Control.Name(), $oCheckBox.Control.Label(), __LOWriter_FormConGetObj($oCheckBox.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oCheckBox.Control.WritingMode(), _
				$oCheckBox.Control.Enabled(), $oCheckBox.Control.EnableVisible(), $oCheckBox.Control.Printable(), $oCheckBox.Control.Tabstop(), $oCheckBox.Control.TabIndex(), _
				$oCheckBox.Control.DefaultState(), __LOWriter_FormConSetGetFontDesc($oCheckBox), $oCheckBox.Control.VisualEffect(), $oCheckBox.Control.Align(), $oCheckBox.Control.VerticalAlign(), _
				$oCheckBox.Control.BackgroundColor(), $oCheckBox.Control.MultiLine(), $oCheckBox.Control.Graphic(), $oCheckBox.Control.ImagePosition(), $oCheckBox.Control.TriState(), _
				$oCheckBox.Control.Tag(), $oCheckBox.Control.HelpText(), $oCheckBox.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCheckBox.Control.Name = $sName
		$iError = ($oCheckBox.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oCheckBox.Control.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCheckBox.Control.Label = $sLabel
		$iError = ($oCheckBox.Control.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oLabelField = Default) Then
		$oCheckBox.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oCheckBox.Control.LabelControl = $oLabelField.Control()
		$iError = ($oCheckBox.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iTxtDir = Default) Then
		$oCheckBox.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oCheckBox.Control.WritingMode = $iTxtDir
		$iError = ($oCheckBox.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oCheckBox.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oCheckBox.Control.Enabled = $bEnabled
		$iError = ($oCheckBox.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bVisible = Default) Then
		$oCheckBox.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oCheckBox.Control.EnableVisible = $bVisible
		$iError = ($oCheckBox.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bPrintable = Default) Then
		$oCheckBox.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oCheckBox.Control.Printable = $bPrintable
		$iError = ($oCheckBox.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bTabStop = Default) Then
		$oCheckBox.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oCheckBox.Control.Tabstop = $bTabStop
		$iError = ($oCheckBox.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 256) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oCheckBox.Control.TabIndex = $iTabOrder
		$iError = ($oCheckBox.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iDefaultState = Default) Then
		$iError = BitOR($iError, 512) ; Can't Default DefaultState.

	ElseIf ($iDefaultState <> Null) Then
		If Not __LO_IntIsBetween($iDefaultState, $LOW_FORM_CON_CHKBX_STATE_NOT_SELECTED, $LOW_FORM_CON_CHKBX_STATE_NOT_DEFINED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oCheckBox.Control.DefaultState = $iDefaultState
		$iError = ($oCheckBox.Control.DefaultState() = $iDefaultState) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		__LOWriter_FormConSetGetFontDesc($oCheckBox, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($iStyle = Default) Then
		$oCheckBox.Control.setPropertyToDefault("VisualEffect")

	ElseIf ($iStyle <> Null) Then
		If Not __LO_IntIsBetween($iStyle, $LOW_FORM_CON_BORDER_3D, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oCheckBox.Control.VisualEffect = $iStyle
		$iError = ($oCheckBox.Control.VisualEffect() = $iStyle) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($iAlign = Default) Then
		$oCheckBox.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oCheckBox.Control.Align = $iAlign
		$iError = ($oCheckBox.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iVertAlign = Default) Then
		$oCheckBox.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oCheckBox.Control.VerticalAlign = $iVertAlign
		$iError = ($oCheckBox.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iBackColor = Default) Then
		$oCheckBox.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oCheckBox.Control.BackgroundColor = $iBackColor
		$iError = ($oCheckBox.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($bWordBreak = Default) Then
		$oCheckBox.Control.setPropertyToDefault("MultiLine")

	ElseIf ($bWordBreak <> Null) Then
		If Not IsBool($bWordBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oCheckBox.Control.MultiLine = $bWordBreak
		$iError = ($oCheckBox.Control.MultiLine() = $bWordBreak) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($sGraphics = Default) Then
		$oCheckBox.Control.setPropertyToDefault("ImageURL")
		$oCheckBox.Control.setPropertyToDefault("Graphic")

	ElseIf ($sGraphics <> Null) Then
		If Not IsString($sGraphics) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oCheckBox.Control.ImageURL = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)
		$iError = ($oCheckBox.Control.ImageURL() = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($iGraphicAlign = Default) Then
		$oCheckBox.Control.setPropertyToDefault("ImagePosition")

	ElseIf ($iGraphicAlign <> Null) Then
		If Not __LO_IntIsBetween($iGraphicAlign, $LOW_FORM_CON_IMG_ALIGN_LEFT_TOP, $LOW_FORM_CON_IMG_ALIGN_CENTERED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oCheckBox.Control.ImagePosition = $iGraphicAlign
		$iError = ($oCheckBox.Control.ImagePosition() = $iGraphicAlign) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($bTriState = Default) Then
		$oCheckBox.Control.setPropertyToDefault("TriState")

	ElseIf ($bTriState <> Null) Then
		If Not IsBool($bTriState) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oCheckBox.Control.TriState = $bTriState
		$iError = ($oCheckBox.Control.TriState = $bTriState) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 524288) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oCheckBox.Control.Tag = $sAddInfo
		$iError = ($oCheckBox.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($sHelpText = Default) Then
		$oCheckBox.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oCheckBox.Control.HelpText = $sHelpText
		$iError = ($oCheckBox.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($sHelpURL = Default) Then
		$oCheckBox.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oCheckBox.Control.HelpURL = $sHelpURL
		$iError = ($oCheckBox.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConCheckBoxGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConCheckBoxState
; Description ...: Set or Retrieve the current Checkbox state.
; Syntax ........: _LOWriter_FormConCheckBoxState(ByRef $oCheckBox[, $iState = Null])
; Parameters ....: $oCheckBox           - [in/out] an object. A Checkbox Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iState              - [optional] an integer value (0-2). Default is Null. The current checked state of the Checkbox, $LOW_FORM_CON_CHKBX_STATE_NOT_DEFINED is only available if $bTriState is True. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCheckBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oCheckBox not a Check Box Control.
;                  @Error 1 @Extended 3 Return 0 = $iState not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current control State.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iState
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current Check Box State as an Integer, matching one of the constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current check box state.
;                  Call $iState with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConCheckBoxGeneral, _LOWriter_FormConCheckBoxData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConCheckBoxState(ByRef $oCheckBox, $iState = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurState

	If Not IsObj($oCheckBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oCheckBox) <> $LOW_FORM_CON_TYPE_CHECK_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iState) Then
		$iCurState = $oCheckBox.Control.State()
		If Not IsInt($iCurState) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurState)
	EndIf

	If ($iState = Default) Then
		$oCheckBox.Control.setPropertyToDefault("State")

	Else
		If Not __LO_IntIsBetween($iState, $LOW_FORM_CON_CHKBX_STATE_NOT_SELECTED, $LOW_FORM_CON_CHKBX_STATE_NOT_DEFINED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCheckBox.Control.State = $iState
		$iError = ($oCheckBox.Control.State() = $iState) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConCheckBoxState

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConComboBoxData
; Description ...: Set or Retrieve Combo Box Data Properties.
; Syntax ........: _LOWriter_FormConComboBoxData(ByRef $oComboBox[, $sDataField = Null[, $bEmptyIsNull = Null[, $bInputRequired = Null[, $iType = Null[, $sListContent = Null]]]]])
; Parameters ....: $oComboBox           - [in/out] an object. A Combo Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bEmptyIsNull        - [optional] a boolean value. Default is Null. If True, an empty string will be treated as a Null value.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
;                  $iType               - [optional] an integer value (1-5). Default is Null. The type of content to fill the control with. See Constants $LOW_FORM_CON_SOURCE_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sListContent        - [optional] a string value. Default is Null. Default is Null. The SQL statement, Table Name, etc., depending on the value of $iType.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComboBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oComboBox not a Combo Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bEmptyIsNull not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bInputRequired not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iType not an Integer, less than 1 or greater than 5. See Constants $LOW_FORM_CON_SOURCE_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $sListContent not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bEmptyIsNull
;                  |                               4 = Error setting $bInputRequired
;                  |                               8 = Error setting $iType
;                  |                               16 = Error setting $sListContent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConComboBoxValue, _LOWriter_FormConComboBoxGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConComboBoxData(ByRef $oComboBox, $sDataField = Null, $bEmptyIsNull = Null, $bInputRequired = Null, $iType = Null, $sListContent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[5]

	If Not IsObj($oComboBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oComboBox) <> $LOW_FORM_CON_TYPE_COMBO_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bEmptyIsNull, $bInputRequired, $iType, $sListContent) Then
		__LO_ArrayFill($avControl, $oComboBox.Control.DataField(), $oComboBox.Control.ConvertEmptyToNull(), $oComboBox.Control.InputRequired(), _
				$oComboBox.Control.ListSourceType(), $oComboBox.Control.ListSource())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oComboBox.Control.DataField = $sDataField
		$iError = ($oComboBox.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bEmptyIsNull <> Null) Then
		If Not IsBool($bEmptyIsNull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oComboBox.Control.ConvertEmptyToNull = $bEmptyIsNull
		$iError = ($oComboBox.Control.ConvertEmptyToNull() = $bEmptyIsNull) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oComboBox.Control.InputRequired = $bInputRequired
		$iError = ($oComboBox.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iType <> Null) Then
		If Not __LO_IntIsBetween($iType, $LOW_FORM_CON_SOURCE_TYPE_TABLE, $LOW_FORM_CON_SOURCE_TYPE_TABLE_FIELDS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oComboBox.Control.ListSourceType = $iType
		$iError = ($oComboBox.Control.ListSourceType() = $iType) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sListContent <> Null) Then
		If Not IsString($sListContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oComboBox.Control.ListSource = $sListContent
		$iError = ($oComboBox.Control.ListSource() = $sListContent) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConComboBoxData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConComboBoxGeneral
; Description ...: Set or Retrieve general Combo Box Properties.
; Syntax ........: _LOWriter_FormConComboBoxGeneral(ByRef $oComboBox[, $sName = Null[, $oLabelField = Null[, $iTxtDir = Null[, $iMaxLen = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $iMouseScroll = Null[, $bTabStop = Null[, $iTabOrder = Null[, $asList = Null[, $sDefaultTxt = Null[, $mFont = Null[, $iAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bDropdown = Null[, $iLines = Null[, $bAutoFill = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oComboBox           - [in/out] an object. A Combo Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iMaxLen             - [optional] an integer value (-1-2147483647). Default is Null. The maximum text length that the Combo box will accept.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $asList              - [optional] an array of strings. Default is Null. An array of entries. See remarks.
;                  $sDefaultTxt         - [optional] a string value. Default is Null. The default text of the combo Box.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bDropdown           - [optional] a boolean value. Default is Null. If True, the Combo Box will behave like a dropdown.
;                  $iLines              - [optional] an integer value. Default is Null. If $bDropdown is True, $iLines specifies how many lines are shown in the dropdown list.
;                  $bAutoFill           - [optional] a boolean value. Default is Null. If True, the Autofill function is enabled.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComboBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oComboBox not a Combo Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 5 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 6 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $iMaxLen not an Integer, less than -1 or greater than 2147483647.
;                  @Error 1 @Extended 8 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 15 Return 0 = $asList not an Array.
;                  @Error 1 @Extended 16 Return ? = Element contained in $asList not a String. Returning problem element position.
;                  @Error 1 @Extended 17 Return 0 = $sDefaultTxt not a String.
;                  @Error 1 @Extended 18 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 19 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 20 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 21 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 22 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215. $bDropdown not a Boolean.
;                  @Error 1 @Extended 23 Return 0 = $bDropdown not a Boolean.
;                  @Error 1 @Extended 24 Return 0 = $iLines not an Integer, less than -2147483648 or greater than 2147483647.
;                  @Error 1 @Extended 25 Return 0 = $bAutoFill not an Boolean.
;                  @Error 1 @Extended 26 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 27 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 28 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 29 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $oLabelField
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $iMaxLen
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bVisible
;                  |                               64 = Error setting $bReadOnly
;                  |                               128 = Error setting $bPrintable
;                  |                               256 = Error setting $iMouseScroll
;                  |                               512 = Error setting $bTabStop
;                  |                               1024 = Error setting $iTabOrder
;                  |                               2048 = Error setting $asList
;                  |                               4096 = Error setting $sDefaultTxt
;                  |                               8192 = Error setting $mFont
;                  |                               16384 = Error setting $iAlign
;                  |                               32768 = Error setting $iBackColor
;                  |                               65536 = Error setting $iBorder
;                  |                               131072 = Error setting $iBorderColor
;                  |                               262144 = Error setting $bDropdown
;                  |                               524288 = Error setting $iLines
;                  |                               1048576 = Error setting $bAutoFill
;                  |                               2097152 = Error setting $bHideSel
;                  |                               4194304 = Error setting $sAddInfo
;                  |                               8388608 = Error setting $sHelpText
;                  |                               16777216 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 25 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $asList, $sDefaultTxt, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormConComboBoxValue, _LOWriter_FormConComboBoxData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConComboBoxGeneral(ByRef $oComboBox, $sName = Null, $oLabelField = Null, $iTxtDir = Null, $iMaxLen = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $iMouseScroll = Null, $bTabStop = Null, $iTabOrder = Null, $asList = Null, $sDefaultTxt = Null, $mFont = Null, $iAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bDropdown = Null, $iLines = Null, $bAutoFill = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[25]

	If Not IsObj($oComboBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oComboBox) <> $LOW_FORM_CON_TYPE_COMBO_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $oLabelField, $iTxtDir, $iMaxLen, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $iMouseScroll, $bTabStop, $iTabOrder, $asList, $sDefaultTxt, $mFont, $iAlign, $iBackColor, $iBorder, $iBorderColor, $bDropdown, $iLines, $bAutoFill, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oComboBox.Control.Name(), __LOWriter_FormConGetObj($oComboBox.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oComboBox.Control.WritingMode(), $oComboBox.Control.MaxTextLen(), _
				$oComboBox.Control.Enabled(), $oComboBox.Control.EnableVisible(), $oComboBox.Control.ReadOnly(), $oComboBox.Control.Printable(), $oComboBox.Control.MouseWheelBehavior(), _
				$oComboBox.Control.Tabstop(), $oComboBox.Control.TabIndex(), $oComboBox.Control.StringItemList(), $oComboBox.Control.DefaultText(), __LOWriter_FormConSetGetFontDesc($oComboBox), _
				$oComboBox.Control.Align(), $oComboBox.Control.BackgroundColor(), $oComboBox.Control.Border(), $oComboBox.Control.BorderColor(), $oComboBox.Control.Dropdown(), _
				$oComboBox.Control.LineCount(), $oComboBox.Control.Autocomplete(), $oComboBox.Control.HideInactiveSelection(), $oComboBox.Control.Tag(), $oComboBox.Control.HelpText(), _
				$oComboBox.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oComboBox.Control.Name = $sName
		$iError = ($oComboBox.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oLabelField = Default) Then
		$oComboBox.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oComboBox.Control.LabelControl = $oLabelField.Control()
		$iError = ($oComboBox.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oComboBox.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oComboBox.Control.WritingMode = $iTxtDir
		$iError = ($oComboBox.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iMaxLen = Default) Then
		$oComboBox.Control.setPropertyToDefault("MaxTextLen")

	ElseIf ($iMaxLen <> Null) Then
		If Not __LO_IntIsBetween($iMaxLen, -1, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oComboBox.Control.MaxTextLen = $iMaxLen
		$iError = ($oComboBox.Control.MaxTextLen() = $iMaxLen) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oComboBox.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oComboBox.Control.Enabled = $bEnabled
		$iError = ($oComboBox.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bVisible = Default) Then
		$oComboBox.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oComboBox.Control.EnableVisible = $bVisible
		$iError = ($oComboBox.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bReadOnly = Default) Then
		$oComboBox.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oComboBox.Control.ReadOnly = $bReadOnly
		$iError = ($oComboBox.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bPrintable = Default) Then
		$oComboBox.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oComboBox.Control.Printable = $bPrintable
		$iError = ($oComboBox.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iMouseScroll = Default) Then
		$oComboBox.Control.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oComboBox.Control.MouseWheelBehavior = $iMouseScroll
		$iError = ($oComboBox.Control.MouseWheelBehavior = $iMouseScroll) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bTabStop = Default) Then
		$oComboBox.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oComboBox.Control.Tabstop = $bTabStop
		$iError = ($oComboBox.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oComboBox.Control.TabIndex = $iTabOrder
		$iError = ($oComboBox.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($asList = Default) Then
		$iError = BitOR($iError, 2048) ; Can't Default StringItemList.

	ElseIf ($asList <> Null) Then
		If Not IsArray($asList) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		For $i = 0 To UBound($asList) - 1
			If Not IsString($asList[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, $i)

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next
		$oComboBox.Control.StringItemList = $asList
		$iError = (UBound($oComboBox.Control.StringItemList()) = UBound($asList)) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($sDefaultTxt = Default) Then
		$iError = BitOR($iError, 4096) ; Can't Default DefaultText.

	ElseIf ($sDefaultTxt <> Null) Then
		If Not IsString($sDefaultTxt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oComboBox.Control.DefaultText = $sDefaultTxt
		$iError = ($oComboBox.Control.DefaultText() = $sDefaultTxt) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 8192) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		__LOWriter_FormConSetGetFontDesc($oComboBox, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iAlign = Default) Then
		$oComboBox.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oComboBox.Control.Align = $iAlign
		$iError = ($oComboBox.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iBackColor = Default) Then
		$oComboBox.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oComboBox.Control.BackgroundColor = $iBackColor
		$iError = ($oComboBox.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($iBorder = Default) Then
		$oComboBox.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oComboBox.Control.Border = $iBorder
		$iError = ($oComboBox.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($iBorderColor = Default) Then
		$oComboBox.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oComboBox.Control.BorderColor = $iBorderColor
		$iError = ($oComboBox.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($bDropdown = Default) Then
		$oComboBox.Control.setPropertyToDefault("Dropdown")

	ElseIf ($bDropdown <> Null) Then
		If Not IsBool($bDropdown) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oComboBox.Control.Dropdown = $bDropdown
		$iError = ($oComboBox.Control.Dropdown() = $bDropdown) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($iLines = Default) Then
		$oComboBox.Control.setPropertyToDefault("LineCount")

	ElseIf ($iLines <> Null) Then
		If Not __LO_IntIsBetween($iLines, -2147483648, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oComboBox.Control.LineCount = $iLines
		$iError = ($oComboBox.Control.LineCount = $iLines) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($bAutoFill = Default) Then
		$oComboBox.Control.setPropertyToDefault("Autocomplete")

	ElseIf ($bAutoFill <> Null) Then
		If Not IsBool($bAutoFill) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oComboBox.Control.Autocomplete = $bAutoFill
		$iError = ($oComboBox.Control.Autocomplete() = $bAutoFill) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($bHideSel = Default) Then
		$oComboBox.Control.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, 0)

		$oComboBox.Control.HideInactiveSelection = $bHideSel
		$iError = ($oComboBox.Control.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 4194304) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 27, 0)

		$oComboBox.Control.Tag = $sAddInfo
		$iError = ($oComboBox.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	If ($sHelpText = Default) Then
		$oComboBox.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 28, 0)

		$oComboBox.Control.HelpText = $sHelpText
		$iError = ($oComboBox.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 8388608))
	EndIf

	If ($sHelpURL = Default) Then
		$oComboBox.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 29, 0)

		$oComboBox.Control.HelpURL = $sHelpURL
		$iError = ($oComboBox.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 16777216))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConComboBoxGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConComboBoxValue
; Description ...: Set or Retrieve a Combo Box's current selection.
; Syntax ........: _LOWriter_FormConComboBoxValue(ByRef $oComboBox[, $sValue = Null])
; Parameters ....: $oComboBox           - [in/out] an object. A Combo Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sValue              - [optional] a string value. Default is Null. The current value in the Combo Box's entry field. Value doesn't need to match an available field.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComboBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oComboBox not a Combo Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sValue not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve currently selected value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sValue
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning currently selected Combo Box value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the currently selected value.
;                  Call $sValue with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConComboBoxGeneral, _LOWriter_FormConComboBoxData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConComboBoxValue(ByRef $oComboBox, $sValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $sCurValue

	If Not IsObj($oComboBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oComboBox) <> $LOW_FORM_CON_TYPE_COMBO_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sValue) Then
		$sCurValue = $oComboBox.Control.CurrentValue()
		If Not IsString($sCurValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurValue)
	EndIf

	If ($sValue = Default) Then
		$oComboBox.Control.setPropertyToDefault("Text")

	Else
		If Not IsString($sValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oComboBox.Control.Text = $sValue
		$iError = ($oComboBox.Control.Text() = $sValue) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConComboBoxValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConCurrencyFieldData
; Description ...: Set or Retrieve Currency Field Data Properties.
; Syntax ........: _LOWriter_FormConCurrencyFieldData(ByRef $oCurrencyField[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oCurrencyField      - [in/out] an object. A Currency Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCurrencyField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oCurrencyField not a Currency Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConCurrencyFieldValue, _LOWriter_FormConCurrencyFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConCurrencyFieldData(ByRef $oCurrencyField, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oCurrencyField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oCurrencyField) <> $LOW_FORM_CON_TYPE_CURRENCY_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oCurrencyField.Control.DataField(), $oCurrencyField.Control.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCurrencyField.Control.DataField = $sDataField
		$iError = ($oCurrencyField.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCurrencyField.Control.InputRequired = $bInputRequired
		$iError = ($oCurrencyField.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConCurrencyFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConCurrencyFieldGeneral
; Description ...: Set or Retrieve general Currency Field properties.
; Syntax ........: _LOWriter_FormConCurrencyFieldGeneral(ByRef $oCurrencyField[, $sName = Null[, $oLabelField = Null[, $iTxtDir = Null[, $bStrict = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $iMouseScroll = Null[, $bTabStop = Null[, $iTabOrder = Null[, $nMin = Null[, $nMax = Null[, $iIncr = Null[, $nDefault = Null[, $iDecimal = Null[, $bThousandsSep = Null[, $sCurrSymbol = Null[, $bPrefix = Null[, $bSpin = Null[, $bRepeat = Null[, $iDelay = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oCurrencyField      - [in/out] an object. A Currency Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bStrict             - [optional] a boolean value. Default is Null. If True, strict formatting is enabled.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $nMin                - [optional] a general number value. Default is Null. The minimum value the control can be set to.
;                  $nMax                - [optional] a general number value. Default is Null. The maximum value the control can be set to.
;                  $iIncr               - [optional] an integer value. Default is Null. The amount to Increase or Decrease the value by.
;                  $nDefault            - [optional] a general number value. Default is Null. The default value the control will be set to.
;                  $iDecimal            - [optional] an integer value (0-20). Default is Null. The amount of decimal accuracy.
;                  $bThousandsSep       - [optional] a boolean value. Default is Null. If True, a thousands separator will be added.
;                  $sCurrSymbol         - [optional] a string value. Default is Null. The symbol to use for currency.
;                  $bPrefix             - [optional] a boolean value. Default is Null. If True, the currency symbol is prefixed to the value.
;                  $bSpin               - [optional] a boolean value. Default is Null. If True, the field will act as a spin button.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCurrencyField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oCurrencyField not a Currency Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 5 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 6 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bStrict not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 15 Return 0 = $nMin not a Number.
;                  @Error 1 @Extended 16 Return 0 = $nMax not a Number.
;                  @Error 1 @Extended 17 Return 0 = $iIncr not an Integer.
;                  @Error 1 @Extended 18 Return 0 = $nDefault not a Number.
;                  @Error 1 @Extended 19 Return 0 = $iDecimal not an Integer, less than 0 or greater than 20.
;                  @Error 1 @Extended 20 Return 0 = $bThousandsSep not a Boolean.
;                  @Error 1 @Extended 21 Return 0 = $sCurrSymbol not a String.
;                  @Error 1 @Extended 22 Return 0 = $bPrefix not a Boolean.
;                  @Error 1 @Extended 23 Return 0 = $bSpin not a Boolean.
;                  @Error 1 @Extended 24 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 25 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 26 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 27 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 28 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 29 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 30 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 31 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 32 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 33 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 34 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 35 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $oLabelField
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bStrict
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bVisible
;                  |                               64 = Error setting $bReadOnly
;                  |                               128 = Error setting $bPrintable
;                  |                               256 = Error setting $iMouseScroll
;                  |                               512 = Error setting $bTabStop
;                  |                               1024 = Error setting $iTabOrder
;                  |                               2048 = Error setting $nMin
;                  |                               4096 = Error setting $nMax
;                  |                               8192 = Error setting $iIncr
;                  |                               16384 = Error setting $nDefault
;                  |                               32768 = Error setting $iDecimal
;                  |                               65536 = Error setting $bThousandsSep
;                  |                               131072 = Error setting $sCurrSymbol
;                  |                               262144 = Error setting $bPrefix
;                  |                               524288 = Error setting $bSpin
;                  |                               1048576 = Error setting $bRepeat
;                  |                               2097152 = Error setting $iDelay
;                  |                               4194304 = Error setting $mFont
;                  |                               8388608 = Error setting $iAlign
;                  |                               16777216 = Error setting $iVertAlign
;                  |                               33554432 = Error setting $iBackColor
;                  |                               67108864 = Error setting $iBorder
;                  |                               134217728 = Error setting $iBorderColor
;                  |                               268435456 = Error setting $bHideSel
;                  |                               536870912 = Error setting $sAddInfo
;                  |                               1073741824 = Error setting $sHelpText
;                  |                               -1 = Error setting $sHelpURL (See remarks.)
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 32 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  If there is an error setting $sHelpURL, the @Extended value for Property setting error will be either -1, or if there are other errors present, a negative value of the error value.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormConCurrencyFieldValue, _LOWriter_FormConCurrencyFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConCurrencyFieldGeneral(ByRef $oCurrencyField, $sName = Null, $oLabelField = Null, $iTxtDir = Null, $bStrict = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $iMouseScroll = Null, $bTabStop = Null, $iTabOrder = Null, $nMin = Null, $nMax = Null, $iIncr = Null, $nDefault = Null, $iDecimal = Null, $bThousandsSep = Null, $sCurrSymbol = Null, $bPrefix = Null, $bSpin = Null, $bRepeat = Null, $iDelay = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[32]

	If Not IsObj($oCurrencyField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oCurrencyField) <> $LOW_FORM_CON_TYPE_CURRENCY_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $oLabelField, $iTxtDir, $bStrict, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $iMouseScroll, $bTabStop, $iTabOrder, $nMin, $nMax, $iIncr, $nDefault, $iDecimal, $bThousandsSep, $sCurrSymbol, $bPrefix, $bSpin, $bRepeat, $iDelay, $mFont, $iAlign, $iVertAlign, $iBackColor, $iBorder, $iBorderColor, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oCurrencyField.Control.Name(), __LOWriter_FormConGetObj($oCurrencyField.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oCurrencyField.Control.WritingMode(), _
				$oCurrencyField.Control.StrictFormat(), $oCurrencyField.Control.Enabled(), $oCurrencyField.Control.EnableVisible(), $oCurrencyField.Control.ReadOnly(), _
				$oCurrencyField.Control.Printable(), $oCurrencyField.Control.MouseWheelBehavior(), $oCurrencyField.Control.Tabstop(), $oCurrencyField.Control.TabIndex(), _
				$oCurrencyField.Control.ValueMin(), $oCurrencyField.Control.ValueMax(), $oCurrencyField.Control.ValueStep(), $oCurrencyField.Control.DefaultValue(), _
				$oCurrencyField.Control.DecimalAccuracy(), $oCurrencyField.Control.ShowThousandsSeparator(), $oCurrencyField.Control.CurrencySymbol(), _
				$oCurrencyField.Control.PrependCurrencySymbol(), $oCurrencyField.Control.Spin(), $oCurrencyField.Control.Repeat(), $oCurrencyField.Control.RepeatDelay(), _
				__LOWriter_FormConSetGetFontDesc($oCurrencyField), $oCurrencyField.Control.Align(), $oCurrencyField.Control.VerticalAlign(), $oCurrencyField.Control.BackgroundColor(), _
				$oCurrencyField.Control.Border(), $oCurrencyField.Control.BorderColor(), $oCurrencyField.Control.HideInactiveSelection(), $oCurrencyField.Control.Tag(), _
				$oCurrencyField.Control.HelpText(), $oCurrencyField.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCurrencyField.Control.Name = $sName
		$iError = ($oCurrencyField.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oLabelField = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oCurrencyField.Control.LabelControl = $oLabelField.Control()
		$iError = ($oCurrencyField.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oCurrencyField.Control.WritingMode = $iTxtDir
		$iError = ($oCurrencyField.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bStrict = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("StrictFormat")

	ElseIf ($bStrict <> Null) Then
		If Not IsBool($bStrict) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oCurrencyField.Control.StrictFormat = $bStrict
		$iError = ($oCurrencyField.Control.StrictFormat() = $bStrict) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oCurrencyField.Control.Enabled = $bEnabled
		$iError = ($oCurrencyField.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bVisible = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oCurrencyField.Control.EnableVisible = $bVisible
		$iError = ($oCurrencyField.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bReadOnly = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oCurrencyField.Control.ReadOnly = $bReadOnly
		$iError = ($oCurrencyField.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bPrintable = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oCurrencyField.Control.Printable = $bPrintable
		$iError = ($oCurrencyField.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iMouseScroll = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oCurrencyField.Control.MouseWheelBehavior = $iMouseScroll
		$iError = ($oCurrencyField.Control.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bTabStop = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oCurrencyField.Control.Tabstop = $bTabStop
		$iError = ($oCurrencyField.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oCurrencyField.Control.TabIndex = $iTabOrder
		$iError = ($oCurrencyField.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($nMin = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("ValueMin")

	ElseIf ($nMin <> Null) Then
		If Not IsNumber($nMin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oCurrencyField.Control.ValueMin = $nMin
		$iError = ($oCurrencyField.Control.ValueMin() = $nMin) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($nMax = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("ValueMax")

	ElseIf ($nMax <> Null) Then
		If Not IsNumber($nMax) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oCurrencyField.Control.ValueMax = $nMax
		$iError = ($oCurrencyField.Control.ValueMax() = $nMax) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iIncr = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("ValueStep")

	ElseIf ($iIncr <> Null) Then
		If Not IsInt($iIncr) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oCurrencyField.Control.ValueStep = $iIncr
		$iError = ($oCurrencyField.Control.ValueStep() = $iIncr) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($nDefault = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("DefaultValue")

	ElseIf ($nDefault <> Null) Then
		If Not IsNumber($nDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oCurrencyField.Control.DefaultValue = $nDefault
		$iError = ($oCurrencyField.Control.DefaultValue() = $nDefault) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iDecimal = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("DecimalAccuracy")

	ElseIf ($iDecimal <> Null) Then
		If Not __LO_IntIsBetween($iDecimal, 0, 20) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oCurrencyField.Control.DecimalAccuracy = $iDecimal
		$iError = ($oCurrencyField.Control.DecimalAccuracy() = $iDecimal) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bThousandsSep = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("ShowThousandsSeparator")

	ElseIf ($bThousandsSep <> Null) Then
		If Not IsBool($bThousandsSep) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oCurrencyField.Control.ShowThousandsSeparator = $bThousandsSep
		$iError = ($oCurrencyField.Control.ShowThousandsSeparator() = $bThousandsSep) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($sCurrSymbol = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("CurrencySymbol")

	ElseIf ($sCurrSymbol <> Null) Then
		If Not IsString($sCurrSymbol) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oCurrencyField.Control.CurrencySymbol = $sCurrSymbol
		$iError = ($oCurrencyField.Control.CurrencySymbol() = $sCurrSymbol) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($bPrefix = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("PrependCurrencySymbol")

	ElseIf ($bPrefix <> Null) Then
		If Not IsBool($bPrefix) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oCurrencyField.Control.PrependCurrencySymbol = $bPrefix
		$iError = ($oCurrencyField.Control.PrependCurrencySymbol() = $bPrefix) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($bSpin = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("Spin")

	ElseIf ($bSpin <> Null) Then
		If Not IsBool($bSpin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oCurrencyField.Control.Spin = $bSpin
		$iError = ($oCurrencyField.Control.Spin() = $bSpin) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($bRepeat = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oCurrencyField.Control.Repeat = $bRepeat
		$iError = ($oCurrencyField.Control.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($iDelay = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oCurrencyField.Control.RepeatDelay = $iDelay
		$iError = ($oCurrencyField.Control.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 4194304) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, 0)

		__LOWriter_FormConSetGetFontDesc($oCurrencyField, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	If ($iAlign = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 27, 0)

		$oCurrencyField.Control.Align = $iAlign
		$iError = ($oCurrencyField.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 8388608))
	EndIf

	If ($iVertAlign = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 28, 0)

		$oCurrencyField.Control.VerticalAlign = $iVertAlign
		$iError = ($oCurrencyField.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 16777216))
	EndIf

	If ($iBackColor = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 29, 0)

		$oCurrencyField.Control.BackgroundColor = $iBackColor
		$iError = ($oCurrencyField.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 33554432))
	EndIf

	If ($iBorder = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 30, 0)

		$oCurrencyField.Control.Border = $iBorder
		$iError = ($oCurrencyField.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 67108864))
	EndIf

	If ($iBorderColor = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 31, 0)

		$oCurrencyField.Control.BorderColor = $iBorderColor
		$iError = ($oCurrencyField.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 134217728))
	EndIf

	If ($bHideSel = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 32, 0)

		$oCurrencyField.Control.HideInactiveSelection = $bHideSel
		$iError = ($oCurrencyField.Control.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 268435456))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 536870912) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 33, 0)

		$oCurrencyField.Control.Tag = $sAddInfo
		$iError = ($oCurrencyField.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 536870912))
	EndIf

	If ($sHelpText = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 34, 0)

		$oCurrencyField.Control.HelpText = $sHelpText
		$iError = ($oCurrencyField.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 1073741824))
	EndIf

	If ($sHelpURL = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 35, 0)

		$oCurrencyField.Control.HelpURL = $sHelpURL
		$iError = ($oCurrencyField.Control.HelpURL() = $sHelpURL) ? ($iError) : (($iError > 0) ? ($iError * -1) : (BitOR($iError, -1)))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConCurrencyFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConCurrencyFieldValue
; Description ...: Set or retrieve the current Currency field value.
; Syntax ........: _LOWriter_FormConCurrencyFieldValue(ByRef $oCurrencyField[, $nValue = Null])
; Parameters ....: $oCurrencyField      - [in/out] an object. A Currency Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $nValue              - [optional] a general number value. Default is Null. The value to set the field to.
; Return values .: Success: 1 or Number
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCurrencyField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oCurrencyField not a Currency Field Control.
;                  @Error 1 @Extended 3 Return 0 = $nValue not a Number.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $nValue
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Number = Success. All optional parameters were called with Null, returning current value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current value. Return will be Null if a value hasn't been set.
;                  Call $nValue with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConCurrencyFieldGeneral, _LOWriter_FormConCurrencyFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConCurrencyFieldValue(ByRef $oCurrencyField, $nValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $nCurVal

	If Not IsObj($oCurrencyField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oCurrencyField) <> $LOW_FORM_CON_TYPE_CURRENCY_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($nValue) Then
		$nCurVal = $oCurrencyField.Control.Value() ; Value is Null when not set.

		Return SetError($__LO_STATUS_SUCCESS, 1, $nCurVal)
	EndIf

	If ($nValue = Default) Then
		$oCurrencyField.Control.setPropertyToDefault("Value")

	Else
		If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCurrencyField.Control.Value = $nValue
		$iError = ($oCurrencyField.Control.Value() = $nValue) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConCurrencyFieldValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConDateFieldData
; Description ...: Set or Retrieve Date Field Data Properties.
; Syntax ........: _LOWriter_FormConDateFieldData(ByRef $oDateField[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oDateField          - [in/out] an object. A Date Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDateField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oDateField not a Date Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConDateFieldValue, _LOWriter_FormConDateFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConDateFieldData(ByRef $oDateField, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oDateField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oDateField) <> $LOW_FORM_CON_TYPE_DATE_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oDateField.Control.DataField(), $oDateField.Control.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDateField.Control.DataField = $sDataField
		$iError = ($oDateField.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDateField.Control.InputRequired = $bInputRequired
		$iError = ($oDateField.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConDateFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConDateFieldGeneral
; Description ...: Set or Retrieve general Date Field properties.
; Syntax ........: _LOWriter_FormConDateFieldGeneral(ByRef $oDateField[, $sName = Null[, $oLabelField = Null[, $iTxtDir = Null[, $bStrict = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $iMouseScroll = Null[, $bTabStop = Null[, $iTabOrder = Null[, $tDateMin = Null[, $tDateMax = Null[, $iFormat = Null[, $tDateDefault = Null[, $bSpin = Null[, $bRepeat = Null[, $iDelay = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bDropdown = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oDateField          - [in/out] an object. A Date Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bStrict             - [optional] a boolean value. Default is Null. If True, strict formatting is enabled.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $tDateMin            - [optional] a dll struct value. Default is Null. The minimum date allowed to be entered, created previously by _LOWriter_DateStructCreate.
;                  $tDateMax            - [optional] a dll struct value. Default is Null. The maximum date allowed to be entered, created previously by _LOWriter_DateStructCreate.
;                  $iFormat             - [optional] an integer value (0-11). Default is Null. The Date Format to display the content in. See Constants $LOW_FORM_CON_DATE_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $tDateDefault        - [optional] a dll struct value. Default is Null. The Default date to display, created previously by _LOWriter_DateStructCreate.
;                  $bSpin               - [optional] a boolean value. Default is Null. If True, the field will act as a spin button.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bDropdown           - [optional] a boolean value. Default is Null. If True, the field will behave as a dropdown control.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDateField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oDateField not a Date Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 5 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 6 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bStrict not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 15 Return 0 = $tDateMin not an Object.
;                  @Error 1 @Extended 16 Return 0 = $tDateMax not an Object.
;                  @Error 1 @Extended 17 Return 0 = $iFormat not an Integer, less then 0 or greater than 11. See Constants $LOW_FORM_CON_DATE_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 18 Return 0 = $tDateDefault not an Object.
;                  @Error 1 @Extended 19 Return 0 = $bSpin not a Boolean.
;                  @Error 1 @Extended 20 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 21 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 22 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 23 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 24 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 25 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 26 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 27 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 28 Return 0 = $bDropdown not a Boolean.
;                  @Error 1 @Extended 29 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 30 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 31 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 32 Return 0 = $sHelpURL not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.DateTime" Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.util.Date" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current Minimum Date.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve current Maximum Date.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $oLabelField
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bStrict
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bVisible
;                  |                               64 = Error setting $bReadOnly
;                  |                               128 = Error setting $bPrintable
;                  |                               256 = Error setting $iMouseScroll
;                  |                               512 = Error setting $bTabStop
;                  |                               1024 = Error setting $iTabOrder
;                  |                               2048 = Error setting $tDateMin
;                  |                               4096 = Error setting $tDateMax
;                  |                               8192 = Error setting $iFormat
;                  |                               16384 = Error setting $tDateDefault
;                  |                               32768 = Error setting $bSpin
;                  |                               65536 = Error setting $bRepeat
;                  |                               131072 = Error setting $iDelay
;                  |                               262144 = Error setting $mFont
;                  |                               524288 = Error setting $iAlign
;                  |                               1048576 = Error setting $iVertAlign
;                  |                               2097152 = Error setting $iBackColor
;                  |                               4194304 = Error setting $iBorder
;                  |                               8388608 = Error setting $iBorderColor
;                  |                               16777216 = Error setting $bDropdown
;                  |                               33554432 = Error setting $bHideSel
;                  |                               67108864 = Error setting $sAddInfo
;                  |                               134217728 = Error setting $sHelpText
;                  |                               268435456 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 29 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormConDateFieldValue, _LOWriter_FormConDateFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConDateFieldGeneral(ByRef $oDateField, $sName = Null, $oLabelField = Null, $iTxtDir = Null, $bStrict = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $iMouseScroll = Null, $bTabStop = Null, $iTabOrder = Null, $tDateMin = Null, $tDateMax = Null, $iFormat = Null, $tDateDefault = Null, $bSpin = Null, $bRepeat = Null, $iDelay = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bDropdown = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tDate, $tCurMin, $tCurMax, $tCurDefault
	Local $avControl[29]

	If Not IsObj($oDateField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oDateField) <> $LOW_FORM_CON_TYPE_DATE_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $oLabelField, $iTxtDir, $bStrict, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $iMouseScroll, $bTabStop, $iTabOrder, $tDateMin, $tDateMax, $iFormat, $tDateDefault, $bSpin, $bRepeat, $iDelay, $mFont, $iAlign, $iVertAlign, $iBackColor, $iBorder, $iBorderColor, $bDropdown, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		$tDate = $oDateField.Control.DateMin()
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$tCurMin = __LO_CreateStruct("com.sun.star.util.DateTime")
		If Not IsObj($tCurMin) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tCurMin.Year = $tDate.Year()
		$tCurMin.Month = $tDate.Month()
		$tCurMin.Day = $tDate.Day()

		$tDate = $oDateField.Control.DateMax()
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$tCurMax = __LO_CreateStruct("com.sun.star.util.DateTime")
		If Not IsObj($tCurMax) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tCurMax.Year = $tDate.Year()
		$tCurMax.Month = $tDate.Month()
		$tCurMax.Day = $tDate.Day()

		$tDate = $oDateField.Control.DefaultDate() ; Default date is Null when not set.
		If IsObj($tDate) Then
			$tCurDefault = __LO_CreateStruct("com.sun.star.util.DateTime")
			If Not IsObj($tCurDefault) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tCurDefault.Year = $tDate.Year()
			$tCurDefault.Month = $tDate.Month()
			$tCurDefault.Day = $tDate.Day()

		Else
			$tCurDefault = $tDate
		EndIf

		__LO_ArrayFill($avControl, $oDateField.Control.Name(), __LOWriter_FormConGetObj($oDateField.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oDateField.Control.WritingMode(), $oDateField.Control.StrictFormat(), _
				$oDateField.Control.Enabled(), $oDateField.Control.EnableVisible(), $oDateField.Control.ReadOnly(), $oDateField.Control.Printable(), $oDateField.Control.MouseWheelBehavior(), _
				$oDateField.Control.Tabstop(), $oDateField.Control.TabIndex(), $tCurMin, $tCurMax, $oDateField.Control.DateFormat(), $tCurDefault, $oDateField.Control.Spin(), _
				$oDateField.Control.Repeat(), $oDateField.Control.RepeatDelay(), __LOWriter_FormConSetGetFontDesc($oDateField), $oDateField.Control.Align(), $oDateField.Control.VerticalAlign(), _
				$oDateField.Control.BackgroundColor(), $oDateField.Control.Border(), $oDateField.Control.BorderColor(), $oDateField.Control.Dropdown(), _
				$oDateField.Control.HideInactiveSelection(), $oDateField.Control.Tag(), $oDateField.Control.HelpText(), $oDateField.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDateField.Control.Name = $sName
		$iError = ($oDateField.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oLabelField = Default) Then
		$oDateField.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oDateField.Control.LabelControl = $oLabelField.Control()
		$iError = ($oDateField.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oDateField.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oDateField.Control.WritingMode = $iTxtDir
		$iError = ($oDateField.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bStrict = Default) Then
		$oDateField.Control.setPropertyToDefault("StrictFormat")

	ElseIf ($bStrict <> Null) Then
		If Not IsBool($bStrict) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oDateField.Control.StrictFormat = $bStrict
		$iError = ($oDateField.Control.StrictFormat() = $bStrict) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oDateField.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oDateField.Control.Enabled = $bEnabled
		$iError = ($oDateField.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bVisible = Default) Then
		$oDateField.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oDateField.Control.EnableVisible = $bVisible
		$iError = ($oDateField.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bReadOnly = Default) Then
		$oDateField.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oDateField.Control.ReadOnly = $bReadOnly
		$iError = ($oDateField.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bPrintable = Default) Then
		$oDateField.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oDateField.Control.Printable = $bPrintable
		$iError = ($oDateField.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iMouseScroll = Default) Then
		$oDateField.Control.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oDateField.Control.MouseWheelBehavior = $iMouseScroll
		$iError = ($oDateField.Control.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bTabStop = Default) Then
		$oDateField.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oDateField.Control.Tabstop = $bTabStop
		$iError = ($oDateField.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oDateField.Control.TabIndex = $iTabOrder
		$iError = ($oDateField.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($tDateMin = Default) Then
		$oDateField.Control.setPropertyToDefault("DateMin")

	ElseIf ($tDateMin <> Null) Then
		If Not IsObj($tDateMin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$tDate = __LO_CreateStruct("com.sun.star.util.Date")
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tDate.Year = $tDateMin.Year()
		$tDate.Month = $tDateMin.Month()
		$tDate.Day = $tDateMin.Day()

		$oDateField.Control.DateMin = $tDate
		$iError = (__LOWriter_DateStructCompare($oDateField.Control.DateMin(), $tDate, True)) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($tDateMax = Default) Then
		$oDateField.Control.setPropertyToDefault("DateMax")

	ElseIf ($tDateMax <> Null) Then
		If Not IsObj($tDateMax) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$tDate = __LO_CreateStruct("com.sun.star.util.Date")
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tDate.Year = $tDateMax.Year()
		$tDate.Month = $tDateMax.Month()
		$tDate.Day = $tDateMax.Day()

		$oDateField.Control.DateMax = $tDate
		$iError = (__LOWriter_DateStructCompare($oDateField.Control.DateMax(), $tDate, True)) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iFormat = Default) Then
		$oDateField.Control.setPropertyToDefault("DateFormat")

	ElseIf ($iFormat <> Null) Then
		If Not __LO_IntIsBetween($iFormat, $LOW_FORM_CON_DATE_FRMT_SYSTEM_SHORT, $LOW_FORM_CON_DATE_FRMT_SHORT_YYYY_MM_DD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oDateField.Control.DateFormat = $iFormat
		$iError = ($oDateField.Control.DateFormat() = $iFormat) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($tDateDefault = Default) Then
		$oDateField.Control.setPropertyToDefault("DefaultDate")

	ElseIf ($tDateDefault <> Null) Then
		If Not IsObj($tDateDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$tDate = __LO_CreateStruct("com.sun.star.util.Date")
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tDate.Year = $tDateDefault.Year()
		$tDate.Month = $tDateDefault.Month()
		$tDate.Day = $tDateDefault.Day()

		$oDateField.Control.DefaultDate = $tDate
		$iError = (__LOWriter_DateStructCompare($oDateField.Control.DefaultDate(), $tDate, True)) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($bSpin = Default) Then
		$oDateField.Control.setPropertyToDefault("Spin")

	ElseIf ($bSpin <> Null) Then
		If Not IsBool($bSpin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oDateField.Control.Spin = $bSpin
		$iError = ($oDateField.Control.Spin() = $bSpin) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bRepeat = Default) Then
		$oDateField.Control.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oDateField.Control.Repeat = $bRepeat
		$iError = ($oDateField.Control.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($iDelay = Default) Then
		$oDateField.Control.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oDateField.Control.RepeatDelay = $iDelay
		$iError = ($oDateField.Control.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 262144) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		__LOWriter_FormConSetGetFontDesc($oDateField, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($iAlign = Default) Then
		$oDateField.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oDateField.Control.Align = $iAlign
		$iError = ($oDateField.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($iVertAlign = Default) Then
		$oDateField.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oDateField.Control.VerticalAlign = $iVertAlign
		$iError = ($oDateField.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($iBackColor = Default) Then
		$oDateField.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oDateField.Control.BackgroundColor = $iBackColor
		$iError = ($oDateField.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($iBorder = Default) Then
		$oDateField.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, 0)

		$oDateField.Control.Border = $iBorder
		$iError = ($oDateField.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	If ($iBorderColor = Default) Then
		$oDateField.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 27, 0)

		$oDateField.Control.BorderColor = $iBorderColor
		$iError = ($oDateField.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 8388608))
	EndIf

	If ($bDropdown = Default) Then
		$oDateField.Control.setPropertyToDefault("Dropdown")

	ElseIf ($bDropdown <> Null) Then
		If Not IsBool($bDropdown) Then Return SetError($__LO_STATUS_INPUT_ERROR, 28, 0)

		$oDateField.Control.Dropdown = $bDropdown
		$iError = ($oDateField.Control.Dropdown() = $bDropdown) ? ($iError) : (BitOR($iError, 16777216))
	EndIf

	If ($bHideSel = Default) Then
		$oDateField.Control.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 29, 0)

		$oDateField.Control.HideInactiveSelection = $bHideSel
		$iError = ($oDateField.Control.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 33554432))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 67108864) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 30, 0)

		$oDateField.Control.Tag = $sAddInfo
		$iError = ($oDateField.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 67108864))
	EndIf

	If ($sHelpText = Default) Then
		$oDateField.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 31, 0)

		$oDateField.Control.HelpText = $sHelpText
		$iError = ($oDateField.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 134217728))
	EndIf

	If ($sHelpURL = Default) Then
		$oDateField.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 32, 0)

		$oDateField.Control.HelpURL = $sHelpURL
		$iError = ($oDateField.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 268435456))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConDateFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConDateFieldValue
; Description ...: Set or retrieve the current Date field value.
; Syntax ........: _LOWriter_FormConDateFieldValue(ByRef $oDateField[, $tDateValue = Null])
; Parameters ....: $oDateField          - [in/out] an object. A Date Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $tDateValue          - [optional] a dll struct value. Default is Null. The date to set the field to, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Structure
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDateField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oDateField not a Date Field Control.
;                  @Error 1 @Extended 3 Return 0 = $tDateValue not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.DateTime" Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.util.Date" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $tDateValue
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Structure = Success. All optional parameters were called with Null, returning current Date value as a Date Structure.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current Date value. Return will be Null if the Date hasn't been set.
;                  Call $tDateValue with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConDateFieldGeneral, _LOWriter_FormConDateFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConDateFieldValue(ByRef $oDateField, $tDateValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tDate, $tCurDate

	If Not IsObj($oDateField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oDateField) <> $LOW_FORM_CON_TYPE_DATE_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($tDateValue) Then
		$tDate = $oDateField.Control.Date() ; Date is Null when not set.
		If IsObj($tDate) Then
			$tCurDate = __LO_CreateStruct("com.sun.star.util.DateTime")
			If Not IsObj($tCurDate) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tCurDate.Year = $tDate.Year()
			$tCurDate.Month = $tDate.Month()
			$tCurDate.Day = $tDate.Day()

		Else
			$tCurDate = $tDate
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $tCurDate)
	EndIf

	If ($tDateValue = Default) Then
		$oDateField.Control.setPropertyToDefault("Date")

	Else
		If Not IsObj($tDateValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tDate = __LO_CreateStruct("com.sun.star.util.Date")
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tDate.Year = $tDateValue.Year()
		$tDate.Month = $tDateValue.Month()
		$tDate.Day = $tDateValue.Day()

		$oDateField.Control.Date = $tDate
		$iError = (__LOWriter_DateStructCompare($oDateField.Control.Date(), $tDate, True)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConDateFieldValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConDelete
; Description ...: Delete a Form Control or Control from a Group.
; Syntax ........: _LOWriter_FormConDelete(ByRef $oControl)
; Parameters ....: $oControl            - [in/out] an object. A Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Control's parent Object.
;                  @Error 3 @Extended 2 Return 0 = Cannot delete the last control in a Grouped control.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Control was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You cannot delete the last control contained in a Grouped Control.
; Related .......: _LOWriter_FormConInsert, _LOWriter_FormConsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConDelete(ByRef $oControl)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oParent

	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oParent = $oControl.Parent() ; Retrieve the parent.
	If Not IsObj($oParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($oParent.supportsService("com.sun.star.drawing.GroupShape") And ($oParent.Count() = 1)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oParent.remove($oControl)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FormConDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConFileSelFieldGeneral
; Description ...: Set or Retrieve general File Selection Field properties.
; Syntax ........: _LOWriter_FormConFileSelFieldGeneral(ByRef $oFileSel[, $sName = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $bTabStop = Null[, $iTabOrder = Null[, $sDefaultTxt = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oFileSel            - [in/out] an object. A File Selection Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $sDefaultTxt         - [optional] a string value. Default is Null. The default text to display in the field.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFileSel not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oFileSel not a File Selection Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 11 Return 0 = $sDefaultTxt not a String.
;                  @Error 1 @Extended 12 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 13 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 14 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 15 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 16 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 17 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 18 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 19 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 20 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 21 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $iTxtDir
;                  |                               4 = Error setting $bEnabled
;                  |                               8 = Error setting $bVisible
;                  |                               16 = Error setting $bReadOnly
;                  |                               32 = Error setting $bPrintable
;                  |                               64 = Error setting $bTabStop
;                  |                               128 = Error setting $iTabOrder
;                  |                               256 = Error setting $sDefaultTxt
;                  |                               512 = Error setting $mFont
;                  |                               1024 = Error setting $iAlign
;                  |                               2048 = Error setting $iVertAlign
;                  |                               4096 = Error setting $iBackColor
;                  |                               8192 = Error setting $iBorder
;                  |                               16384 = Error setting $iBorderColor
;                  |                               32768 = Error setting $bHideSel
;                  |                               65536 = Error setting $sAddInfo
;                  |                               131072 = Error setting $sHelpText
;                  |                               262144 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 19 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $sDefaultTxt, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormConFileSelFieldValue
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConFileSelFieldGeneral(ByRef $oFileSel, $sName = Null, $iTxtDir = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $bTabStop = Null, $iTabOrder = Null, $sDefaultTxt = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[19]

	If Not IsObj($oFileSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oFileSel) <> $LOW_FORM_CON_TYPE_FILE_SELECTION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $iTxtDir, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $bTabStop, $iTabOrder, $sDefaultTxt, $mFont, $iAlign, $iVertAlign, $iBackColor, $iBorder, $iBorderColor, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oFileSel.Control.Name(), $oFileSel.Control.WritingMode(), $oFileSel.Control.Enabled(), $oFileSel.Control.EnableVisible(), _
				$oFileSel.Control.ReadOnly(), $oFileSel.Control.Printable(), $oFileSel.Control.Tabstop(), $oFileSel.Control.TabIndex(), $oFileSel.Control.DefaultText(), _
				__LOWriter_FormConSetGetFontDesc($oFileSel), $oFileSel.Control.Align(), $oFileSel.Control.VerticalAlign(), $oFileSel.Control.BackgroundColor(), $oFileSel.Control.Border(), _
				$oFileSel.Control.BorderColor(), $oFileSel.Control.HideInactiveSelection(), $oFileSel.Control.Tag(), $oFileSel.Control.HelpText(), $oFileSel.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFileSel.Control.Name = $sName
		$iError = ($oFileSel.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTxtDir = Default) Then
		$oFileSel.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFileSel.Control.WritingMode = $iTxtDir
		$iError = ($oFileSel.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bEnabled = Default) Then
		$oFileSel.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFileSel.Control.Enabled = $bEnabled
		$iError = ($oFileSel.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bVisible = Default) Then
		$oFileSel.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFileSel.Control.EnableVisible = $bVisible
		$iError = ($oFileSel.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bReadOnly = Default) Then
		$oFileSel.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFileSel.Control.ReadOnly = $bReadOnly
		$iError = ($oFileSel.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bPrintable = Default) Then
		$oFileSel.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFileSel.Control.Printable = $bPrintable
		$iError = ($oFileSel.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bTabStop = Default) Then
		$oFileSel.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oFileSel.Control.Tabstop = $bTabStop
		$iError = ($oFileSel.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 128) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oFileSel.Control.TabIndex = $iTabOrder
		$iError = ($oFileSel.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($sDefaultTxt = Default) Then
		$iError = BitOR($iError, 256) ; Can't Default DefaultText.

	ElseIf ($sDefaultTxt <> Null) Then
		If Not IsString($sDefaultTxt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oFileSel.Control.DefaultText = $sDefaultTxt
		$iError = ($oFileSel.Control.DefaultText() = $sDefaultTxt) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 512) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		__LOWriter_FormConSetGetFontDesc($oFileSel, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iAlign = Default) Then
		$oFileSel.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oFileSel.Control.Align = $iAlign
		$iError = ($oFileSel.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($iVertAlign = Default) Then
		$oFileSel.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oFileSel.Control.VerticalAlign = $iVertAlign
		$iError = ($oFileSel.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($iBackColor = Default) Then
		$oFileSel.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oFileSel.Control.BackgroundColor = $iBackColor
		$iError = ($oFileSel.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iBorder = Default) Then
		$oFileSel.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oFileSel.Control.Border = $iBorder
		$iError = ($oFileSel.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iBorderColor = Default) Then
		$oFileSel.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oFileSel.Control.BorderColor = $iBorderColor
		$iError = ($oFileSel.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($bHideSel = Default) Then
		$oFileSel.Control.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oFileSel.Control.HideInactiveSelection = $bHideSel
		$iError = ($oFileSel.Control.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 65536) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oFileSel.Control.Tag = $sAddInfo
		$iError = ($oFileSel.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($sHelpText = Default) Then
		$oFileSel.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oFileSel.Control.HelpText = $sHelpText
		$iError = ($oFileSel.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($sHelpURL = Default) Then
		$oFileSel.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oFileSel.Control.HelpURL = $sHelpURL
		$iError = ($oFileSel.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConFileSelFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConFileSelFieldValue
; Description ...: Set or retrieve the current File Selection Field value.
; Syntax ........: _LOWriter_FormConFileSelFieldValue(ByRef $oFileSel[, $sValue = Null])
; Parameters ....: $oFileSel            - [in/out] an object. A File Selection Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sValue              - [optional] a string value. Default is Null. The value to set the field to.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFileSel not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oFileSel not a File Selection Control.
;                  @Error 1 @Extended 3 Return 0 = $sValue not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the current value of the control.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sValue
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning current value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current value.
;                  Call $sValue with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConFileSelFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConFileSelFieldValue(ByRef $oFileSel, $sValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $sCurValue

	If Not IsObj($oFileSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oFileSel) <> $LOW_FORM_CON_TYPE_FILE_SELECTION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sValue) Then
		$sCurValue = $oFileSel.Control.Text()
		If Not IsString($sCurValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurValue)
	EndIf

	If ($sValue = Default) Then
		$oFileSel.Control.setPropertyToDefault("Text")

	Else
		If Not IsString($sValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFileSel.Control.Text = $sValue
		$iError = ($oFileSel.Control.Text() = $sValue) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConFileSelFieldValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConFormattedFieldData
; Description ...: Set or Retrieve Formatted Field Data Properties.
; Syntax ........: _LOWriter_FormConFormattedFieldData(ByRef $oFormatField[, $sDataField = Null[, $bEmptyIsNull = Null[, $bInputRequired = Null[, $bFilter = Null]]]])
; Parameters ....: $oFormatField        - [in/out] an object. A Formatted Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bEmptyIsNull        - [optional] a boolean value. Default is Null. If True, an empty string will be treated as a Null value.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
;                  $bFilter             - [optional] a boolean value. Default is Null. If True, filter proposal is active.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormatField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oFormatField not a Formatted Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bEmptyIsNull not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bInputRequired not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bFilter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bEmptyIsNull
;                  |                               4 = Error setting $bInputRequired
;                  |                               8 = Error setting $bFilter
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConFormattedFieldValue, _LOWriter_FormConFormattedFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConFormattedFieldData(ByRef $oFormatField, $sDataField = Null, $bEmptyIsNull = Null, $bInputRequired = Null, $bFilter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[4]

	If Not IsObj($oFormatField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oFormatField) <> $LOW_FORM_CON_TYPE_FORMATTED_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bEmptyIsNull, $bInputRequired, $bFilter) Then
		__LO_ArrayFill($avControl, $oFormatField.Control.DataField(), $oFormatField.Control.ConvertEmptyToNull(), $oFormatField.Control.InputRequired(), _
				$oFormatField.Control.UseFilterValueProposal())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFormatField.Control.DataField = $sDataField
		$iError = ($oFormatField.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bEmptyIsNull <> Null) Then
		If Not IsBool($bEmptyIsNull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFormatField.Control.ConvertEmptyToNull = $bEmptyIsNull
		$iError = ($oFormatField.Control.ConvertEmptyToNull() = $bEmptyIsNull) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFormatField.Control.InputRequired = $bInputRequired
		$iError = ($oFormatField.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bFilter <> Null) Then
		If Not IsBool($bFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFormatField.Control.UseFilterValueProposal = $bFilter
		$iError = ($oFormatField.Control.UseFilterValueProposal() = $bFilter) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConFormattedFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConFormattedFieldGeneral
; Description ...: Set or Retrieve general Formatted Field properties.
; Syntax ........: _LOWriter_FormConFormattedFieldGeneral(ByRef $oFormatField[, $sName = Null[, $oLabelField = Null[, $iTxtDir = Null[, $iMaxLen = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $iMouseScroll = Null[, $bTabStop = Null[, $iTabOrder = Null[, $nMin = Null[, $nMax = Null[, $nDefault = Null[, $iFormat = Null[, $bSpin = Null[, $bRepeat = Null[, $iDelay = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oFormatField        - [in/out] an object. A Formatted Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iMaxLen             - [optional] an integer value (-1-2147483647). Default is Null. The maximum text length that the Formatted field will accept.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $nMin                - [optional] a general number value. Default is Null. The minimum value allowed in the field.
;                  $nMax                - [optional] a general number value. Default is Null. The maximum value allowed in the field.
;                  $nDefault            - [optional] a general number value. Default is Null. The default value of the field.
;                  $iFormat             - [optional] an integer value. Default is Null. The Number Format Key to display the content in, retrieved from a previous _LOWriter_FormatKeysGetList call, or created by _LOWriter_FormatKeyCreate function.
;                  $bSpin               - [optional] a boolean value. Default is Null. If True, the field will act as a spin button.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormatField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oFormatField not a Formatted Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 5 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 6 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $iMaxLen not an Integer, less than -1 or greater than 2147483647.
;                  @Error 1 @Extended 8 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 15 Return 0 = $nMin not a Number.
;                  @Error 1 @Extended 16 Return 0 = $nMax not a Number.
;                  @Error 1 @Extended 17 Return 0 = $nDefault not a Number.
;                  @Error 1 @Extended 18 Return 0 = $iFormat not an Integer.
;                  @Error 1 @Extended 19 Return 0 = Format key called in $iFormat not found in document.
;                  @Error 1 @Extended 20 Return 0 = $bSpin not a Boolean.
;                  @Error 1 @Extended 21 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 22 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 23 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 24 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 25 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 26 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 27 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 28 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 29 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 30 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 31 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 32 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $oLabelField
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $iMaxLen
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bVisible
;                  |                               64 = Error setting $bReadOnly
;                  |                               128 = Error setting $bPrintable
;                  |                               256 = Error setting $iMouseScroll
;                  |                               512 = Error setting $bTabStop
;                  |                               1024 = Error setting $iTabOrder
;                  |                               2048 = Error setting $nMin
;                  |                               4096 = Error setting $nMax
;                  |                               8192 = Error setting $nDefault
;                  |                               16384 = Error setting $iFormat
;                  |                               32768 = Error setting $bSpin
;                  |                               65536 = Error setting $bRepeat
;                  |                               131072 = Error setting $iDelay
;                  |                               262144 = Error setting $mFont
;                  |                               524288 = Error setting $iAlign
;                  |                               1048576 = Error setting $iVertAlign
;                  |                               2097152 = Error setting $iBackColor
;                  |                               4194304 = Error setting $iBorder
;                  |                               8388608 = Error setting $iBorderColor
;                  |                               16777216 = Error setting $bHideSel
;                  |                               33554432 = Error setting $sAddInfo
;                  |                               67108864 = Error setting $sHelpText
;                  |                               134217728 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 28 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormatKeyCreate, _LOWriter_FormatKeysGetList, _LOWriter_FormConFormattedFieldValue, _LOWriter_FormConFormattedFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConFormattedFieldGeneral(ByRef $oFormatField, $sName = Null, $oLabelField = Null, $iTxtDir = Null, $iMaxLen = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $iMouseScroll = Null, $bTabStop = Null, $iTabOrder = Null, $nMin = Null, $nMax = Null, $nDefault = Null, $iFormat = Null, $bSpin = Null, $bRepeat = Null, $iDelay = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oDoc
	Local $avControl[28]

	If Not IsObj($oFormatField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oFormatField) <> $LOW_FORM_CON_TYPE_FORMATTED_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $oLabelField, $iTxtDir, $iMaxLen, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $iMouseScroll, $bTabStop, $iTabOrder, $nMin, $nMax, $nDefault, $iFormat, $bSpin, $bRepeat, $iDelay, $mFont, $iAlign, $iVertAlign, $iBackColor, $iBorder, $iBorderColor, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oFormatField.Control.Name(), __LOWriter_FormConGetObj($oFormatField.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oFormatField.Control.WritingMode(), $oFormatField.Control.MaxTextLen(), _
				$oFormatField.Control.Enabled(), $oFormatField.Control.EnableVisible(), $oFormatField.Control.ReadOnly(), $oFormatField.Control.Printable(), _
				$oFormatField.Control.MouseWheelBehavior(), $oFormatField.Control.Tabstop(), $oFormatField.Control.TabIndex(), $oFormatField.Control.EffectiveMin(), _
				$oFormatField.Control.EffectiveMax(), $oFormatField.Control.EffectiveDefault(), $oFormatField.Control.FormatKey(), $oFormatField.Control.Spin(), _
				$oFormatField.Control.Repeat(), $oFormatField.Control.RepeatDelay(), __LOWriter_FormConSetGetFontDesc($oFormatField), $oFormatField.Control.Align(), _
				$oFormatField.Control.VerticalAlign(), $oFormatField.Control.BackgroundColor(), $oFormatField.Control.Border(), $oFormatField.Control.BorderColor(), _
				$oFormatField.Control.HideInactiveSelection(), $oFormatField.Control.Tag(), $oFormatField.Control.HelpText(), $oFormatField.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFormatField.Control.Name = $sName
		$iError = ($oFormatField.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oLabelField = Default) Then
		$oFormatField.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFormatField.Control.LabelControl = $oLabelField.Control()
		$iError = ($oFormatField.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oFormatField.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFormatField.Control.WritingMode = $iTxtDir
		$iError = ($oFormatField.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iMaxLen = Default) Then
		$oFormatField.Control.setPropertyToDefault("MaxTextLen")

	ElseIf ($iMaxLen <> Null) Then
		If Not __LO_IntIsBetween($iMaxLen, -1, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFormatField.Control.MaxTextLen = $iMaxLen
		$iError = ($oFormatField.Control.MaxTextLen = $iMaxLen) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oFormatField.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFormatField.Control.Enabled = $bEnabled
		$iError = ($oFormatField.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bVisible = Default) Then
		$oFormatField.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oFormatField.Control.EnableVisible = $bVisible
		$iError = ($oFormatField.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bReadOnly = Default) Then
		$oFormatField.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oFormatField.Control.ReadOnly = $bReadOnly
		$iError = ($oFormatField.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bPrintable = Default) Then
		$oFormatField.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oFormatField.Control.Printable = $bPrintable
		$iError = ($oFormatField.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iMouseScroll = Default) Then
		$oFormatField.Control.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oFormatField.Control.MouseWheelBehavior = $iMouseScroll
		$iError = ($oFormatField.Control.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bTabStop = Default) Then
		$oFormatField.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oFormatField.Control.Tabstop = $bTabStop
		$iError = ($oFormatField.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oFormatField.Control.TabIndex = $iTabOrder
		$iError = ($oFormatField.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($nMin = Default) Then
		$oFormatField.Control.setPropertyToDefault("EffectiveMin")

	ElseIf ($nMin <> Null) Then
		If Not IsNumber($nMin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oFormatField.Control.EffectiveMin = $nMin
		$iError = ($oFormatField.Control.EffectiveMin() = $nMin) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($nMax = Default) Then
		$oFormatField.Control.setPropertyToDefault("EffectiveMax")

	ElseIf ($nMax <> Null) Then
		If Not IsNumber($nMax) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oFormatField.Control.EffectiveMax = $nMax
		$iError = ($oFormatField.Control.EffectiveMax() = $nMax) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($nDefault = Default) Then
		$oFormatField.Control.setPropertyToDefault("EffectiveDefault")

	ElseIf ($nDefault <> Null) Then
		If Not IsNumber($nDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oFormatField.Control.EffectiveDefault = $nDefault
		$iError = ($oFormatField.Control.EffectiveDefault() = $nDefault) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iFormat = Default) Then
		$oFormatField.Control.setPropertyToDefault("FormatKey")

	ElseIf ($iFormat <> Null) Then
		If Not IsInt($iFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oDoc = $oFormatField.Control.Parent() ; Identify the parent document.
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Do
			$oDoc = $oDoc.getParent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Until $oDoc.supportsService("com.sun.star.text.TextDocument")
		If Not _LOWriter_FormatKeyExists($oDoc, $iFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oFormatField.Control.FormatKey = $iFormat
		$iError = ($oFormatField.Control.FormatKey() = $iFormat) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($bSpin = Default) Then
		$oFormatField.Control.setPropertyToDefault("Spin")

	ElseIf ($bSpin <> Null) Then
		If Not IsBool($bSpin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oFormatField.Control.Spin = $bSpin
		$iError = ($oFormatField.Control.Spin() = $bSpin) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bRepeat = Default) Then
		$oFormatField.Control.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oFormatField.Control.Repeat = $bRepeat
		$iError = ($oFormatField.Control.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($iDelay = Default) Then
		$oFormatField.Control.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oFormatField.Control.RepeatDelay = $iDelay
		$iError = ($oFormatField.Control.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 262144) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		__LOWriter_FormConSetGetFontDesc($oFormatField, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($iAlign = Default) Then
		$oFormatField.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oFormatField.Control.Align = $iAlign
		$iError = ($oFormatField.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($iVertAlign = Default) Then
		$oFormatField.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oFormatField.Control.VerticalAlign = $iVertAlign
		$iError = ($oFormatField.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($iBackColor = Default) Then
		$oFormatField.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, 0)

		$oFormatField.Control.BackgroundColor = $iBackColor
		$iError = ($oFormatField.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($iBorder = Default) Then
		$oFormatField.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 27, 0)

		$oFormatField.Control.Border = $iBorder
		$iError = ($oFormatField.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	If ($iBorderColor = Default) Then
		$oFormatField.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 28, 0)

		$oFormatField.Control.BorderColor = $iBorderColor
		$iError = ($oFormatField.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 8388608))
	EndIf

	If ($bHideSel = Default) Then
		$oFormatField.Control.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 29, 0)

		$oFormatField.Control.HideInactiveSelection = $bHideSel
		$iError = ($oFormatField.Control.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 16777216))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 33554432) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 30, 0)

		$oFormatField.Control.Tag = $sAddInfo
		$iError = ($oFormatField.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 33554432))
	EndIf

	If ($sHelpText = Default) Then
		$oFormatField.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 31, 0)

		$oFormatField.Control.HelpText = $sHelpText
		$iError = ($oFormatField.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 67108864))
	EndIf

	If ($sHelpURL = Default) Then
		$oFormatField.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 32, 0)

		$oFormatField.Control.HelpURL = $sHelpURL
		$iError = ($oFormatField.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 134217728))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConFormattedFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConFormattedFieldValue
; Description ...: Set or Retrieve the current Formatted Field value.
; Syntax ........: _LOWriter_FormConFormattedFieldValue(ByRef $oFormatField[, $nValue = Null])
; Parameters ....: $oFormatField        - [in/out] an object. A Formatted Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $nValue              - [optional] a general number value. Default is Null. The Value to set the Formatted Field to.
; Return values .: Success: 1 or Number
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormatField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oFormatField not a Formatted Field Control.
;                  @Error 1 @Extended 3 Return 0 = $nValue not a Number.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $nValue
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Number = Success. All optional parameters were called with Null, returning current value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current value.
;                  Call $nValue with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConFormattedFieldGeneral, _LOWriter_FormConFormattedFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConFormattedFieldValue(ByRef $oFormatField, $nValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurValue

	If Not IsObj($oFormatField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oFormatField) <> $LOW_FORM_CON_TYPE_FORMATTED_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($nValue) Then
		$iCurValue = $oFormatField.Control.EffectiveValue()
		If Not IsNumber($iCurValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurValue)
	EndIf

	If ($nValue = Default) Then
		$oFormatField.Control.setPropertyToDefault("EffectiveValue")

	Else
		If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFormatField.Control.EffectiveValue = $nValue
		$iError = ($oFormatField.Control.EffectiveValue() = $nValue) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConFormattedFieldValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConGetParent
; Description ...: Retrieve the Parent Form of the called Control.
; Syntax ........: _LOWriter_FormConGetParent(ByRef $oControl)
; Parameters ....: $oControl            - [in/out] an object. A Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oControl not an Control Object and not a Grouped Control.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve parent form Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the Form Object that contains the Control.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Unfortunately I am unable to successfully set the parent for controls. It sets, but doesn't literally "move" the control to the new form, and also causes the control to no-longer work.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConGetParent(ByRef $oControl)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oOldParent

	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If $oControl.supportsService("com.sun.star.drawing.ControlShape") Then
		$oOldParent = $oControl.Control.Parent()

	ElseIf $oControl.supportsService("com.sun.star.drawing.GroupShape") Then     ; If shape is a grouped control
		For $i = 0 To $oControl.Count() - 1
			If $oControl.getByIndex($i).supportsService("com.sun.star.drawing.ControlShape") Then ; Retrieve the controls contained in the grouped control, to find the parent form.
				$oOldParent = $oControl.getByIndex($i).Control.Parent()
				If IsObj($oOldParent) Then ExitLoop
			EndIf
		Next

		If Not IsObj($oOldParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oOldParent)
EndFunc   ;==>_LOWriter_FormConGetParent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConGroupBoxGeneral
; Description ...: Set or Retrieve general GroupBox control properties.
; Syntax ........: _LOWriter_FormConGroupBoxGeneral(ByRef $oGroupBox[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bVisible = Null[, $bPrintable = Null[, $mFont = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]])
; Parameters ....: $oGroupBox           - [in/out] an object. A Groupbox Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oGroupBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oGroupBox not a GroupBox Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 10 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 11 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 12 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bEnabled
;                  |                               16 = Error setting $bVisible
;                  |                               32 = Error setting $bPrintable
;                  |                               64 = Error setting $mFont
;                  |                               128 = Error setting $sAddInfo
;                  |                               256 = Error setting $sHelpText
;                  |                               512 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 10 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $mFont, $sAddInfo.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConGroupBoxGeneral(ByRef $oGroupBox, $sName = Null, $sLabel = Null, $iTxtDir = Null, $bEnabled = Null, $bVisible = Null, $bPrintable = Null, $mFont = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[10]

	If Not IsObj($oGroupBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oGroupBox) <> $LOW_FORM_CON_TYPE_GROUP_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $bEnabled, $bVisible, $bPrintable, $mFont, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oGroupBox.Control.Name(), $oGroupBox.Control.Label(), $oGroupBox.Control.WritingMode(), $oGroupBox.Control.Enabled(), _
				$oGroupBox.Control.EnableVisible(), $oGroupBox.Control.Printable(), __LOWriter_FormConSetGetFontDesc($oGroupBox), $oGroupBox.Control.Tag(), _
				$oGroupBox.Control.HelpText(), $oGroupBox.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oGroupBox.Control.Name = $sName
		$iError = ($oGroupBox.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oGroupBox.Control.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oGroupBox.Control.Label = $sLabel
		$iError = ($oGroupBox.Control.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oGroupBox.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oGroupBox.Control.WritingMode = $iTxtDir
		$iError = ($oGroupBox.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bEnabled = Default) Then
		$oGroupBox.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oGroupBox.Control.Enabled = $bEnabled
		$iError = ($oGroupBox.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bVisible = Default) Then
		$oGroupBox.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oGroupBox.Control.EnableVisible = $bVisible
		$iError = ($oGroupBox.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bPrintable = Default) Then
		$oGroupBox.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oGroupBox.Control.Printable = $bPrintable
		$iError = ($oGroupBox.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 64) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		__LOWriter_FormConSetGetFontDesc($oGroupBox, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 128) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oGroupBox.Control.Tag = $sAddInfo
		$iError = ($oGroupBox.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($sHelpText = Default) Then
		$oGroupBox.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oGroupBox.Control.HelpText = $sHelpText
		$iError = ($oGroupBox.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($sHelpURL = Default) Then
		$oGroupBox.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oGroupBox.Control.HelpURL = $sHelpURL
		$iError = ($oGroupBox.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 512))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConGroupBoxGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConImageButtonGeneral
; Description ...: Set or Retrieve general Image Button properties.
; Syntax ........: _LOWriter_FormConImageButtonGeneral(ByRef $oImageButton[, $sName = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bVisible = Null[, $bPrintable = Null[, $bTabStop = Null[, $iTabOrder = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $iAction = Null[, $sURL = Null[, $sFrame = Null[, $sGraphics = Null[, $iScale = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]])
; Parameters ....: $oImageButton        - [in/out] an object. A Image Button Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iAction             - [optional] an integer value (0-12). Default is Null. The action that occurs when the button is pushed. See Constants $LOW_FORM_CON_PUSH_CMD_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sURL                - [optional] a string value. Default is Null. The URL or Document path to open.
;                  $sFrame              - [optional] a string value. Default is Null. The frame to open the URL in. See Constants, $LOW_FRAME_TARGET_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sGraphics           - [optional] a string value. Default is Null. The path to an Image file.
;                  $iScale              - [optional] an integer value (0-2). Default is Null. How to scale the image to fit the button. See Constants $LOW_FORM_CON_IMG_BTN_SCALE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImageButton not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oImageButton not an Image Button Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 10 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 11 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 12 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 13 Return 0 = $iAction not an Integer, less than 0 or greater than 12. See Constants $LOW_FORM_CON_PUSH_CMD_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 14 Return 0 = $sURL not a String.
;                  @Error 1 @Extended 15 Return 0 = $sFrame not a String.
;                  @Error 1 @Extended 16 Return 0 = $sFrame not called with correct constant. See Constants, $LOW_FRAME_TARGET_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 17 Return 0 = $sGraphics not a String.
;                  @Error 1 @Extended 18 Return 0 = $iScale not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_IMG_BTN_SCALE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 19 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 20 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 21 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $iTxtDir
;                  |                               4 = Error setting $bEnabled
;                  |                               8 = Error setting $bVisible
;                  |                               16 = Error setting $bPrintable
;                  |                               32 = Error setting $bTabStop
;                  |                               64 = Error setting $iTabOrder
;                  |                               128 = Error setting $iBackColor
;                  |                               256 = Error setting $iBorder
;                  |                               512 = Error setting $iBorderColor
;                  |                               1024 = Error setting $iAction
;                  |                               2048 = Error setting $sURL
;                  |                               4096 = Error setting $sFrame
;                  |                               8192 = Error setting $sGraphics
;                  |                               16384 = Error setting $iScale
;                  |                               32768 = Error setting $sAddInfo
;                  |                               65536 = Error setting $sHelpText
;                  |                               131072 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 18 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $sGraphics is called with an invalid Graphic URL, graphic is set to Null. The Return for $sGraphics is an Image Object.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $sAddInfo.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConImageButtonGeneral(ByRef $oImageButton, $sName = Null, $iTxtDir = Null, $bEnabled = Null, $bVisible = Null, $bPrintable = Null, $bTabStop = Null, $iTabOrder = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $iAction = Null, $sURL = Null, $sFrame = Null, $sGraphics = Null, $iScale = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iBtnAction
	Local Const $__LOW_PUSH_BTN_CMND_PUSH = 0, $__LOW_PUSH_BTN_CMND_SUBMIT = 1, $__LOW_PUSH_BTN_CMND_RESET = 2, $__LOW_PUSH_BTN_CMND_URL = 3 ; com.sun.star.form.FormButtonType
	Local $avControl[18]
	Local $asActions[13]

	If Not IsObj($oImageButton) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oImageButton) <> $LOW_FORM_CON_TYPE_IMAGE_BUTTON) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asActions[$LOW_FORM_CON_PUSH_CMD_NONE] = ""
	$asActions[$LOW_FORM_CON_PUSH_CMD_SUBMIT_FORM] = ""
	$asActions[$LOW_FORM_CON_PUSH_CMD_RESET_FORM] = ""
	$asActions[$LOW_FORM_CON_PUSH_CMD_OPEN] = ""
	$asActions[$LOW_FORM_CON_PUSH_CMD_FIRST_REC] = ".uno:FormController/moveToFirst"
	$asActions[$LOW_FORM_CON_PUSH_CMD_LAST_REC] = ".uno:FormController/moveToLast"
	$asActions[$LOW_FORM_CON_PUSH_CMD_NEXT_REC] = ".uno:FormController/moveToNext"
	$asActions[$LOW_FORM_CON_PUSH_CMD_PREV_REC] = ".uno:FormController/moveToPrev"
	$asActions[$LOW_FORM_CON_PUSH_CMD_SAVE_REC] = ".uno:FormController/saveRecord"
	$asActions[$LOW_FORM_CON_PUSH_CMD_UNDO] = ".uno:FormController/undoRecord"
	$asActions[$LOW_FORM_CON_PUSH_CMD_NEW_REC] = ".uno:FormController/moveToNew"
	$asActions[$LOW_FORM_CON_PUSH_CMD_DELETE_REC] = ".uno:FormController/deleteRecord"
	$asActions[$LOW_FORM_CON_PUSH_CMD_REFRESH_FORM] = ".uno:FormController/refreshForm"

	Switch $oImageButton.Control.ButtonType()
		Case $__LOW_PUSH_BTN_CMND_PUSH
			$iBtnAction = $LOW_FORM_CON_PUSH_CMD_NONE

		Case $__LOW_PUSH_BTN_CMND_SUBMIT
			$iBtnAction = $LOW_FORM_CON_PUSH_CMD_SUBMIT_FORM

		Case $__LOW_PUSH_BTN_CMND_RESET
			$iBtnAction = $LOW_FORM_CON_PUSH_CMD_RESET_FORM

		Case $__LOW_PUSH_BTN_CMND_URL
			$iBtnAction = $LOW_FORM_CON_PUSH_CMD_OPEN

			If StringInStr($oImageButton.Control.TargetURL(), ".uno:FormController/") Then
				Switch $oImageButton.Control.TargetURL()
					Case ".uno:FormController/moveToFirst"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_FIRST_REC

					Case ".uno:FormController/moveToLast"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_LAST_REC

					Case ".uno:FormController/moveToNext"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_NEXT_REC

					Case ".uno:FormController/moveToPrev"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_PREV_REC

					Case ".uno:FormController/saveRecord"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_SAVE_REC

					Case ".uno:FormController/undoRecord"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_UNDO

					Case ".uno:FormController/moveToNew"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_NEW_REC

					Case ".uno:FormController/deleteRecord"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_DELETE_REC

					Case ".uno:FormController/refreshForm"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_REFRESH_FORM
				EndSwitch
			EndIf
	EndSwitch

	If __LO_VarsAreNull($sName, $iTxtDir, $bEnabled, $bVisible, $bPrintable, $bTabStop, $iTabOrder, $iBackColor, $iBorder, $iBorderColor, $iAction, $sURL, $sFrame, $sGraphics, $iScale, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oImageButton.Control.Name(), $oImageButton.Control.WritingMode(), $oImageButton.Control.Enabled(), $oImageButton.Control.EnableVisible(), _
				$oImageButton.Control.Printable(), $oImageButton.Control.Tabstop(), $oImageButton.Control.TabIndex(), $oImageButton.Control.BackgroundColor(), _
				$oImageButton.Control.Border(), $oImageButton.Control.BorderColor(), $iBtnAction, $oImageButton.Control.TargetURL(), $oImageButton.Control.TargetFrame(), _
				$oImageButton.Control.Graphic(), $oImageButton.Control.ScaleMode(), $oImageButton.Control.Tag(), $oImageButton.Control.HelpText(), $oImageButton.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImageButton.Control.Name = $sName
		$iError = ($oImageButton.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTxtDir = Default) Then
		$oImageButton.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImageButton.Control.WritingMode = $iTxtDir
		$iError = ($oImageButton.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bEnabled = Default) Then
		$oImageButton.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImageButton.Control.Enabled = $bEnabled
		$iError = ($oImageButton.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bVisible = Default) Then
		$oImageButton.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImageButton.Control.EnableVisible = $bVisible
		$iError = ($oImageButton.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bPrintable = Default) Then
		$oImageButton.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oImageButton.Control.Printable = $bPrintable
		$iError = ($oImageButton.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bTabStop = Default) Then
		$oImageButton.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oImageButton.Control.Tabstop = $bTabStop
		$iError = ($oImageButton.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 64) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oImageButton.Control.TabIndex = $iTabOrder
		$iError = ($oImageButton.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iBackColor = Default) Then
		$oImageButton.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oImageButton.Control.BackgroundColor = $iBackColor
		$iError = ($oImageButton.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iBorder = Default) Then
		$oImageButton.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oImageButton.Control.Border = $iBorder
		$iError = ($oImageButton.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iBorderColor = Default) Then
		$oImageButton.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oImageButton.Control.BorderColor = $iBorderColor
		$iError = ($oImageButton.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iAction = Default) Then
		$oImageButton.Control.setPropertyToDefault("ButtonType")

		Switch $iBtnAction
			Case $LOW_FORM_CON_PUSH_CMD_NONE, $LOW_FORM_CON_PUSH_CMD_SUBMIT_FORM, $LOW_FORM_CON_PUSH_CMD_RESET_FORM
				$oImageButton.Control.setPropertyToDefault("TargetURL")

			Case $LOW_FORM_CON_PUSH_CMD_OPEN

			Case $LOW_FORM_CON_PUSH_CMD_FIRST_REC, $LOW_FORM_CON_PUSH_CMD_LAST_REC, $LOW_FORM_CON_PUSH_CMD_NEXT_REC, $LOW_FORM_CON_PUSH_CMD_PREV_REC, _
					$LOW_FORM_CON_PUSH_CMD_SAVE_REC, $LOW_FORM_CON_PUSH_CMD_UNDO, $LOW_FORM_CON_PUSH_CMD_NEW_REC, $LOW_FORM_CON_PUSH_CMD_DELETE_REC, _
					$LOW_FORM_CON_PUSH_CMD_REFRESH_FORM
				$oImageButton.Control.setPropertyToDefault("TargetURL")
		EndSwitch

	ElseIf ($iAction <> Null) Then
		If Not __LO_IntIsBetween($iAction, $LOW_FORM_CON_PUSH_CMD_NONE, $LOW_FORM_CON_PUSH_CMD_REFRESH_FORM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		Switch $iAction
			Case $LOW_FORM_CON_PUSH_CMD_NONE
				$oImageButton.Control.ButtonType = $__LOW_PUSH_BTN_CMND_PUSH
				$sURL = $asActions[$iAction]
				$iError = (($oImageButton.Control.ButtonType() = $__LOW_PUSH_BTN_CMND_PUSH)) ? ($iError) : (BitOR($iError, 1024))

			Case $LOW_FORM_CON_PUSH_CMD_SUBMIT_FORM
				$oImageButton.Control.ButtonType = $__LOW_PUSH_BTN_CMND_SUBMIT
				$sURL = $asActions[$iAction]
				$iError = (($oImageButton.Control.ButtonType() = $__LOW_PUSH_BTN_CMND_SUBMIT)) ? ($iError) : (BitOR($iError, 1024))

			Case $LOW_FORM_CON_PUSH_CMD_RESET_FORM
				$oImageButton.Control.ButtonType = $__LOW_PUSH_BTN_CMND_RESET
				$sURL = $asActions[$iAction]
				$iError = (($oImageButton.Control.ButtonType() = $__LOW_PUSH_BTN_CMND_RESET)) ? ($iError) : (BitOR($iError, 1024))

			Case $LOW_FORM_CON_PUSH_CMD_OPEN
				$oImageButton.Control.ButtonType = $__LOW_PUSH_BTN_CMND_URL
				If ($iBtnAction <> $LOW_FORM_CON_PUSH_CMD_OPEN) And ($sURL = Null) Then $sURL = $asActions[$iAction]
				$iError = (($oImageButton.Control.ButtonType() = $__LOW_PUSH_BTN_CMND_URL)) ? ($iError) : (BitOR($iError, 1024))

			Case $LOW_FORM_CON_PUSH_CMD_FIRST_REC, $LOW_FORM_CON_PUSH_CMD_LAST_REC, $LOW_FORM_CON_PUSH_CMD_NEXT_REC, $LOW_FORM_CON_PUSH_CMD_PREV_REC, _
					$LOW_FORM_CON_PUSH_CMD_SAVE_REC, $LOW_FORM_CON_PUSH_CMD_UNDO, $LOW_FORM_CON_PUSH_CMD_NEW_REC, $LOW_FORM_CON_PUSH_CMD_DELETE_REC, _
					$LOW_FORM_CON_PUSH_CMD_REFRESH_FORM
				$oImageButton.Control.ButtonType = $__LOW_PUSH_BTN_CMND_URL
				$sURL = $asActions[$iAction]
				$iError = (($oImageButton.Control.ButtonType() = $__LOW_PUSH_BTN_CMND_URL)) ? ($iError) : (BitOR($iError, 1024))
		EndSwitch
	EndIf

	If ($sURL = Default) Then
		$oImageButton.Control.setPropertyToDefault("TargetURL")

	ElseIf ($sURL <> Null) Then
		If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oImageButton.Control.TargetURL = $sURL
		$iError = ($oImageButton.Control.TargetURL() = $sURL) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($sFrame = Default) Then
		$oImageButton.Control.setPropertyToDefault("TargetFrame")

	ElseIf ($sFrame <> Null) Then
		If Not IsString($sFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)
		If ($sFrame <> $LOW_FRAME_TARGET_TOP) And _
				($sFrame <> $LOW_FRAME_TARGET_PARENT) And _
				($sFrame <> $LOW_FRAME_TARGET_BLANK) And _
				($sFrame <> $LOW_FRAME_TARGET_SELF) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)
		$oImageButton.Control.TargetFrame = $sFrame
		$iError = ($oImageButton.Control.TargetFrame() = $sFrame) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($sGraphics = Default) Then
		$oImageButton.Control.setPropertyToDefault("ImageURL")
		$oImageButton.Control.setPropertyToDefault("Graphic")

	ElseIf ($sGraphics <> Null) Then
		If Not IsString($sGraphics) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oImageButton.Control.ImageURL = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)
		$iError = ($oImageButton.Control.ImageURL() = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iScale = Default) Then
		$oImageButton.Control.setPropertyToDefault("ScaleMode")

	ElseIf ($iScale <> Null) Then
		If Not __LO_IntIsBetween($iScale, $LOW_FORM_CON_IMG_BTN_SCALE_NONE, $LOW_FORM_CON_IMG_BTN_SCALE_FIT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oImageButton.Control.ScaleMode = $iScale
		$iError = ($oImageButton.Control.ScaleMode() = $iScale) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 32768) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oImageButton.Control.Tag = $sAddInfo
		$iError = ($oImageButton.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($sHelpText = Default) Then
		$oImageButton.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oImageButton.Control.HelpText = $sHelpText
		$iError = ($oImageButton.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($sHelpURL = Default) Then
		$oImageButton.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oImageButton.Control.HelpURL = $sHelpURL
		$iError = ($oImageButton.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConImageButtonGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConImageControlData
; Description ...: Set or Retrieve Image Control Data Properties.
; Syntax ........: _LOWriter_FormConImageControlData(ByRef $oImageControl[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oImageControl       - [in/out] an object. A Image Control Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImageControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oImageControl not a Image Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConImageControlGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConImageControlData(ByRef $oImageControl, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oImageControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oImageControl) <> $LOW_FORM_CON_TYPE_IMAGE_CONTROL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oImageControl.Control.DataField(), $oImageControl.Control.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImageControl.Control.DataField = $sDataField
		$iError = ($oImageControl.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImageControl.Control.InputRequired = $bInputRequired
		$iError = ($oImageControl.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConImageControlData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConImageControlGeneral
; Description ...: Set or retrieve general Image control properties.
; Syntax ........: _LOWriter_FormConImageControlGeneral(ByRef $oImageControl[, $sName = Null[, $oLabelField = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $bTabStop = Null[, $iTabOrder = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $sGraphics = Null[, $iScale = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]])
; Parameters ....: $oImageControl       - [in/out] an object. A Image Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $sGraphics           - [optional] a string value. Default is Null. The path to an Image file.
;                  $iScale              - [optional] an integer value (0-2). Default is Null. How to scale the image to fit the button. See Constants $LOW_FORM_CON_IMG_BTN_SCALE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImageControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oImageControl not an Image Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 5 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 6 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 13 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 14 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 15 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 16 Return 0 = $sGraphics not a String.
;                  @Error 1 @Extended 17 Return 0 = $iScale not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_IMG_BTN_SCALE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 18 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 19 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 20 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $oLabelField
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bEnabled
;                  |                               16 = Error setting $bVisible
;                  |                               32 = Error setting $bReadOnly
;                  |                               64 = Error setting $bPrintable
;                  |                               128 = Error setting $bTabStop
;                  |                               256 = Error setting $iTabOrder
;                  |                               512 = Error setting $iBackColor
;                  |                               1024 = Error setting $iBorder
;                  |                               2048 = Error setting $iBorderColor
;                  |                               4096 = Error setting $sGraphics
;                  |                               8192 = Error setting $iScale
;                  |                               16384 = Error setting $sAddInfo
;                  |                               32768 = Error setting $sHelpText
;                  |                               65536 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 17 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $sGraphics is called with an invalid Graphic URL, graphic is set to Null. The Return for $sGraphics is an Image Object.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormConImageControlData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConImageControlGeneral(ByRef $oImageControl, $sName = Null, $oLabelField = Null, $iTxtDir = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $bTabStop = Null, $iTabOrder = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $sGraphics = Null, $iScale = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[17]

	If Not IsObj($oImageControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oImageControl) <> $LOW_FORM_CON_TYPE_IMAGE_CONTROL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $oLabelField, $iTxtDir, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $bTabStop, $iTabOrder, $iBackColor, $iBorder, $iBorderColor, $sGraphics, $iScale, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oImageControl.Control.Name(), __LOWriter_FormConGetObj($oImageControl.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oImageControl.Control.WritingMode(), $oImageControl.Control.Enabled(), _
				$oImageControl.Control.EnableVisible(), $oImageControl.Control.ReadOnly(), $oImageControl.Control.Printable(), $oImageControl.Control.Tabstop(), _
				$oImageControl.Control.TabIndex(), $oImageControl.Control.BackgroundColor(), $oImageControl.Control.Border(), $oImageControl.Control.BorderColor(), _
				$oImageControl.Control.Graphic(), $oImageControl.Control.ScaleMode(), $oImageControl.Control.Tag(), $oImageControl.Control.HelpText(), $oImageControl.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImageControl.Control.Name = $sName
		$iError = ($oImageControl.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oLabelField = Default) Then
		$oImageControl.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImageControl.Control.LabelControl = $oLabelField.Control()
		$iError = ($oImageControl.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oImageControl.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImageControl.Control.WritingMode = $iTxtDir
		$iError = ($oImageControl.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bEnabled = Default) Then
		$oImageControl.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oImageControl.Control.Enabled = $bEnabled
		$iError = ($oImageControl.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bVisible = Default) Then
		$oImageControl.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oImageControl.Control.EnableVisible = $bVisible
		$iError = ($oImageControl.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bReadOnly = Default) Then
		$oImageControl.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oImageControl.Control.ReadOnly = $bReadOnly
		$iError = ($oImageControl.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bPrintable = Default) Then
		$oImageControl.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oImageControl.Control.Printable = $bPrintable
		$iError = ($oImageControl.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bTabStop = Default) Then
		$oImageControl.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oImageControl.Control.Tabstop = $bTabStop
		$iError = ($oImageControl.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 256) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oImageControl.Control.TabIndex = $iTabOrder
		$iError = ($oImageControl.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iBackColor = Default) Then
		$oImageControl.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oImageControl.Control.BackgroundColor = $iBackColor
		$iError = ($oImageControl.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iBorder = Default) Then
		$oImageControl.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oImageControl.Control.Border = $iBorder
		$iError = ($oImageControl.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($iBorderColor = Default) Then
		$oImageControl.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oImageControl.Control.BorderColor = $iBorderColor
		$iError = ($oImageControl.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($sGraphics = Default) Then
		$oImageControl.Control.setPropertyToDefault("ImageURL")
		$oImageControl.Control.setPropertyToDefault("Graphic")

	ElseIf ($sGraphics <> Null) Then
		If Not IsString($sGraphics) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oImageControl.Control.ImageURL = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)
		$iError = ($oImageControl.Control.ImageURL() = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iScale = Default) Then
		$oImageControl.Control.setPropertyToDefault("ScaleMode")

	ElseIf ($iScale <> Null) Then
		If Not __LO_IntIsBetween($iScale, $LOW_FORM_CON_IMG_BTN_SCALE_NONE, $LOW_FORM_CON_IMG_BTN_SCALE_FIT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oImageControl.Control.ScaleMode = $iScale
		$iError = ($oImageControl.Control.ScaleMode() = $iScale) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 16384) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oImageControl.Control.Tag = $sAddInfo
		$iError = ($oImageControl.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($sHelpText = Default) Then
		$oImageControl.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oImageControl.Control.HelpText = $sHelpText
		$iError = ($oImageControl.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($sHelpURL = Default) Then
		$oImageControl.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oImageControl.Control.HelpURL = $sHelpURL
		$iError = ($oImageControl.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConImageControlGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConInsert
; Description ...: Insert a control into a form and document.
; Syntax ........: _LOWriter_FormConInsert(ByRef $oParentForm, $iControl, $iX, $iY, $iWidth, $iHeight[, $sName = ""])
; Parameters ....: $oParentForm         - [in/out] an object. A Form object returned by a previous _LOWriter_FormAdd, _LOWriter_FormGetObjByIndex or _LOWriter_FormsGetList function.
;                  $iControl            - an integer value (1-524288). The control type to insert. See Constants $LOW_FORM_CON_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iX                  - an integer value. The X Coordinate, in Hundredths of a Millimeter (HMM).
;                  $iY                  - an integer value. The Y Coordinate, in Hundredths of a Millimeter (HMM).
;                  $iWidth              - an integer value. The Width of the control, in Hundredths of a Millimeter (HMM).
;                  $iHeight             - an integer value. The Height of the control, in Hundredths of a Millimeter (HMM).
;                  $sName               - [optional] a string value. Default is "". The name of the control, if called with "", a name is automatically given it.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oParentForm not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oParentForm not a Form Object.
;                  @Error 1 @Extended 3 Return 0 = $iControl not an Integer, less than 1 or greater than 524288. See Constants $LOW_FORM_CON_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $sName not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create the Control.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.drawing.ControlShape" Object.
;                  @Error 2 @Extended 3 Return 0 = Failed to create a "com.sun.star.drawing.GroupShape" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form parent document Object.
;                  @Error 3 @Extended 2 Return 0 = Parent Document is ReadOnly.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Control Service name.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Shape Position Structure.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Shape Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Grouped Control was inserted successfully, returning its object.
;                  @Error 0 @Extended 1 Return Object = Success. Control was inserted successfully, returning its object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When inserting a Grouped Control, a Group box will be automatically created and inserted into it.
;                  I have not found a reliable and working way to add controls to a Group of Controls.
; Related .......: _LOWriter_FormConsGetList, _LOWriter_FormConDelete, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConInsert(ByRef $oParentForm, $iControl, $iX, $iY, $iWidth, $iHeight, $sName = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oControl, $oShape, $oGroupShape, $oDoc
	Local $sControl, $sShapeName
	Local $tPos, $tSize
	Local $iCount = 1

	If Not IsObj($oParentForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParentForm.supportsService("com.sun.star.form.component.Form") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iControl, $LOW_FORM_CON_TYPE_CHECK_BOX, $LOW_FORM_CON_TYPE_TIME_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	If ($sName = "") Then
		While $oParentForm.hasByName("AU3_FORM_CNTRL_" & $iCount)
			$iCount += 1
			Sleep((IsInt(($iCount / $__LOWCONST_SLEEP_DIV))) ? (10) : (0))
		WEnd
		$sName = "AU3_FORM_CNTRL_" & $iCount
	EndIf

	$oDoc = $oParentForm ; Identify the parent document.

	Do
		$oDoc = $oDoc.getParent()
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	Until $oDoc.supportsService("com.sun.star.text.TextDocument")

	If $oDoc.IsReadOnly() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iCount = 0

	Do
		$iCount += 1
		For $i = 0 To $oDoc.DrawPage.Count() - 1
			If ($oDoc.DrawPage.getByIndex($i).Name() = "AU3_FORM_SHAPE_" & $iCount) Then ExitLoop
			Sleep((IsInt(($i / $__LOWCONST_SLEEP_DIV))) ? (10) : (0))
		Next
	Until ($i >= $oDoc.DrawPage.Count())

	$sShapeName = "AU3_FORM_SHAPE_" & $iCount

	$sControl = __LOWriter_FormConIdentify(Null, $iControl)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oControl = $oDoc.createInstance($sControl)
	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.ControlShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oControl.Name = $sName

	Switch $iControl
		Case $LOW_FORM_CON_TYPE_CHECK_BOX
			$oControl.Label = "Check Box"

		Case $LOW_FORM_CON_TYPE_GROUP_BOX, $LOW_FORM_CON_TYPE_GROUPED_CONTROL ; For a grouped control, I first need to create a Group Box, and name it appropriately, then the grouped shape.
			$oControl.Label = "Group Box"

		Case $LOW_FORM_CON_TYPE_LABEL
			$oControl.Label = "Label Field"

		Case $LOW_FORM_CON_TYPE_OPTION_BUTTON
			$oControl.Label = "Option Button"

		Case $LOW_FORM_CON_TYPE_PUSH_BUTTON
			$oControl.Label = "Push Button"
	EndSwitch

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	$oShape.Control = $oControl
	$oShape.Name = $sShapeName

	If ($iControl = $LOW_FORM_CON_TYPE_GROUPED_CONTROL) Then
		$oGroupShape = $oDoc.createInstance("com.sun.star.drawing.GroupShape")
		If Not IsObj($oGroupShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$tSize = $oGroupShape.Size()
		If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

		$tSize.Width = $iWidth
		$tSize.Height = $iHeight

		$oGroupShape.Size = $tSize

		$oDoc.DrawPage.Add($oGroupShape)
		$oGroupShape.Add($oShape)

		$oGroupShape.AnchorType = $LOW_ANCHOR_AT_PARAGRAPH ; Have to set anchor after insertion.

		$tPos = $oGroupShape.Position()
		If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$tPos.X = $iX
		$tPos.Y = $iY

		$oGroupShape.Position = $tPos

		$oGroupShape.getByIndex(0).Control.setParent($oParentForm) ; Have to set parent form after otherwise a COM Error is triggered.

		Return SetError($__LO_STATUS_SUCCESS, 1, $oGroupShape)

	Else ; Non-Group Box Control
		$oDoc.DrawPage.Add($oShape)
		$oShape.Control.SetParent($oParentForm)
		$oShape.AnchorType = $LOW_ANCHOR_AT_PARAGRAPH
	EndIf

	; Have to set Position after insertion.
	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oShape.Position = $tPos

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>_LOWriter_FormConInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConLabelGeneral
; Description ...: Set or Retrieve general Label control settings.
; Syntax ........: _LOWriter_FormConLabelGeneral(ByRef $oLabel[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bVisible = Null[, $bPrintable = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bWordBreak = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]])
; Parameters ....: $oLabel              - [in/out] an object. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The name of the Label control.
;                  $sLabel              - [optional] a string value. Default is Null. The Label of the control.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bWordBreak          - [optional] a boolean value. Default is Null. If True, line breaks are allowed to be used.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oLabel not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oLabel not a Label Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 10 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 12 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 13 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 14 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 15 Return 0 = $bWordBreak not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 17 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 18 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bEnabled
;                  |                               16 = Error setting $bVisible
;                  |                               32 = Error setting $bPrintable
;                  |                               64 = Error setting $mFont
;                  |                               128 = Error setting $iAlign
;                  |                               256 = Error setting $iVertAlign
;                  |                               512 = Error setting $iBackColor
;                  |                               1024 = Error setting $iBorder
;                  |                               2048 = Error setting $iBorderColor
;                  |                               4096 = Error setting $bWordBreak
;                  |                               8192 = Error setting $sAddInfo
;                  |                               16384 = Error setting $sHelpText
;                  |                               32768 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 16 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConLabelGeneral(ByRef $oLabel, $sName = Null, $sLabel = Null, $iTxtDir = Null, $bEnabled = Null, $bVisible = Null, $bPrintable = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bWordBreak = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[16]

	If Not IsObj($oLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oLabel) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $bEnabled, $bVisible, $bPrintable, $mFont, $iAlign, $iVertAlign, $iBackColor, $iBorder, $iBorderColor, $bWordBreak, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oLabel.Control.Name(), $oLabel.Control.Label(), $oLabel.Control.WritingMode(), $oLabel.Control.Enabled(), $oLabel.Control.EnableVisible(), _
				$oLabel.Control.Printable(), __LOWriter_FormConSetGetFontDesc($oLabel), $oLabel.Control.Align(), $oLabel.Control.VerticalAlign(), $oLabel.Control.BackgroundColor(), _
				$oLabel.Control.Border(), $oLabel.Control.BorderColor(), $oLabel.Control.MultiLine(), $oLabel.Control.Tag(), $oLabel.Control.HelpText(), $oLabel.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oLabel.Control.Name = $sName
		$iError = ($oLabel.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oLabel.Control.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oLabel.Control.Label = $sLabel
		$iError = ($oLabel.Control.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oLabel.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oLabel.Control.WritingMode = $iTxtDir
		$iError = ($oLabel.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bEnabled = Default) Then
		$oLabel.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oLabel.Control.Enabled = $bEnabled
		$iError = ($oLabel.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bVisible = Default) Then
		$oLabel.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oLabel.Control.EnableVisible = $bVisible
		$iError = ($oLabel.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bPrintable = Default) Then
		$oLabel.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oLabel.Control.Printable = $bPrintable
		$iError = ($oLabel.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 64) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		__LOWriter_FormConSetGetFontDesc($oLabel, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iAlign = Default) Then
		$oLabel.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oLabel.Control.Align = $iAlign
		$iError = ($oLabel.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iVertAlign = Default) Then
		$oLabel.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oLabel.Control.VerticalAlign = $iVertAlign
		$iError = ($oLabel.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iBackColor = Default) Then
		$oLabel.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oLabel.Control.BackgroundColor = $iBackColor
		$iError = ($oLabel.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iBorder = Default) Then
		$oLabel.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oLabel.Control.Border = $iBorder
		$iError = ($oLabel.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($iBorderColor = Default) Then
		$oLabel.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oLabel.Control.BorderColor = $iBorderColor
		$iError = ($oLabel.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($bWordBreak = Default) Then
		$oLabel.Control.setPropertyToDefault("MultiLine")

	ElseIf ($bWordBreak <> Null) Then
		If Not IsBool($bWordBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oLabel.Control.MultiLine = $bWordBreak
		$iError = ($oLabel.Control.MultiLine() = $bWordBreak) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 8192) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oLabel.Control.Tag = $sAddInfo
		$iError = ($oLabel.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($sHelpText = Default) Then
		$oLabel.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oLabel.Control.HelpText = $sHelpText
		$iError = ($oLabel.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($sHelpURL = Default) Then
		$oLabel.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oLabel.Control.HelpURL = $sHelpURL
		$iError = ($oLabel.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConLabelGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConListBoxData
; Description ...: Set or Retrieve List Box Data Properties.
; Syntax ........: _LOWriter_FormConListBoxData(ByRef $oListBox[, $sDataField = Null[, $bInputRequired = Null[, $iType = Null[, $asListContent = Null[, $iBoundField = Null]]]]])
; Parameters ....: $oListBox            - [in/out] an object. A List Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
;                  $iType               - [optional] an integer value (0-5). Default is Null. The type of content to fill the control with. See Constants $LOW_FORM_CON_SOURCE_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $asListContent       - [optional] an array of strings. Default is Null. A single dimension array. See remarks
;                  $iBoundField         - [optional] an integer value (-1-2147483647). Default is Null. The bound data field of a linked table to display.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oListBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oListBox not a List Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iType not an Integer, less than 0 or greater than 5. See Constants $LOW_FORM_CON_SOURCE_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $asListContent not an Array.
;                  @Error 1 @Extended 7 Return 0 = $iType not set to Valuelist and array called in $asListContent has more than 1 element.
;                  @Error 1 @Extended 8 Return 0 = $iBoundField not an Integer, less than -1 or greater than 2147483647.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  |                               4 = Error setting $iType
;                  |                               8 = Error setting $asListContent
;                  |                               16 = Error setting $iBoundField
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
;                  $asListContent is not error checked for the same content, but only that the set array size is the same.
;                  $asListContent should be a single dimension array with a appropriate value in each element. e.g. If $iType is set to Table, the element will contain a Table name. Or if $iType is set to Value List, each element will contain a list item.
;                  For types other than Value list for $iType, the array sound contain a single element.
; Related .......: _LOWriter_FormConListBoxSelection, _LOWriter_FormConListBoxGetCount, _LOWriter_FormConListBoxGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConListBoxData(ByRef $oListBox, $sDataField = Null, $bInputRequired = Null, $iType = Null, $asListContent = Null, $iBoundField = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[5]

	If Not IsObj($oListBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oListBox) <> $LOW_FORM_CON_TYPE_LIST_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired, $iType, $asListContent, $iBoundField) Then
		__LO_ArrayFill($avControl, $oListBox.Control.DataField(), $oListBox.Control.InputRequired(), $oListBox.Control.ListSourceType(), _
				$oListBox.Control.ListSource(), $oListBox.Control.BoundColumn())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oListBox.Control.DataField = $sDataField
		$iError = ($oListBox.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oListBox.Control.InputRequired = $bInputRequired
		$iError = ($oListBox.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iType <> Null) Then
		If Not __LO_IntIsBetween($iType, $LOW_FORM_CON_SOURCE_TYPE_VALUE_LIST, $LOW_FORM_CON_SOURCE_TYPE_TABLE_FIELDS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oListBox.Control.ListSourceType = $iType
		$iError = ($oListBox.Control.ListSourceType() = $iType) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($asListContent <> Null) Then
		If Not IsArray($asListContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If ($oListBox.Control.ListSourceType() <> $LOW_FORM_CON_SOURCE_TYPE_VALUE_LIST) And (UBound($asListContent) > 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oListBox.Control.ListSource = $asListContent
		$iError = (UBound($oListBox.Control.ListSource()) = UBound($asListContent)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iBoundField <> Null) Then
		If Not __LO_IntIsBetween($iBoundField, -1, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oListBox.Control.BoundColumn = $iBoundField
		$iError = ($oListBox.Control.BoundColumn() = $iBoundField) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConListBoxData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConListBoxGeneral
; Description ...: Set or Retrieve general List box properties.
; Syntax ........: _LOWriter_FormConListBoxGeneral(ByRef $oListBox[, $sName = Null[, $oLabelField = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $iMouseScroll = Null[, $bTabStop = Null[, $iTabOrder = Null[, $asList = Null[, $mFont = Null[, $iAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bDropdown = Null[, $iLines = Null[, $bMultiSel = Null[, $aiDefaultSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oListBox            - [in/out] an object. A List Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $asList              - [optional] an array of strings. Default is Null. An array of List entries. See remarks.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bDropdown           - [optional] a boolean value. Default is Null. If True, the List Box will behave like a dropdown.
;                  $iLines              - [optional] an integer value (-2147483648-2147483647). Default is Null. If $bDropdown is True, $iLines specifies how many lines are shown in the dropdown list.
;                  $bMultiSel           - [optional] a boolean value. Default is Null. If True, more than one selection can be made in a list box.
;                  $aiDefaultSel        - [optional] an array of integers. Default is Null. A single dimension array of selection values. See remarks.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oListBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oListBox not a List Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 5 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 6 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 12 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 13 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 14 Return 0 = $asList not an Array.
;                  @Error 1 @Extended 15 Return ? = Element contained in $asList not a String. Returning problem element position.
;                  @Error 1 @Extended 16 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 17 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 18 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 19 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 20 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 21 Return 0 = $bDropdown not a Boolean.
;                  @Error 1 @Extended 22 Return 0 = $iLines not an Integer, less than -2147483648 or greater than 2147483647.
;                  @Error 1 @Extended 23 Return 0 = $bMultiSel not an Boolean.
;                  @Error 1 @Extended 24 Return 0 = $aiDefaultSel not an Array.
;                  @Error 1 @Extended 25 Return ? = Element contained in $aiDefaultSel not an Integer. Returning problem element position.
;                  @Error 1 @Extended 26 Return ? = Integer contained in Element of $aiDefaultSel greater than number of List items. Returning problem element position.
;                  @Error 1 @Extended 27 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 28 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 29 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $oLabelField
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bEnabled
;                  |                               16 = Error setting $bVisible
;                  |                               32 = Error setting $bReadOnly
;                  |                               64 = Error setting $bPrintable
;                  |                               128 = Error setting $iMouseScroll
;                  |                               256 = Error setting $bTabStop
;                  |                               512 = Error setting $iTabOrder
;                  |                               1024 = Error setting $asList
;                  |                               2048 = Error setting $mFont
;                  |                               4096 = Error setting $iAlign
;                  |                               8192 = Error setting $iBackColor
;                  |                               16384 = Error setting $iBorder
;                  |                               32768 = Error setting $iBorderColor
;                  |                               65536 = Error setting $bDropdown
;                  |                               131072 = Error setting $iLines
;                  |                               262144 = Error setting $bMultiSel
;                  |                               524288 = Error setting $aiDefaultSel
;                  |                               1048576 = Error setting $sAddInfo
;                  |                               2097152 = Error setting $sHelpText
;                  |                               4194304 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 23 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The array called for $asList should be a single dimension array, with one List entry as a String, per array element.
;                  The array called for $aiDefaultSel should be a single dimension array, with one Integer value, corresponding to the position in the $asList array, per array element, to indicate which value(s) is/are default.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $asList, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormConListBoxSelection, _LOWriter_FormConListBoxGetCount, _LOWriter_FormConListBoxData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConListBoxGeneral(ByRef $oListBox, $sName = Null, $oLabelField = Null, $iTxtDir = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $iMouseScroll = Null, $bTabStop = Null, $iTabOrder = Null, $asList = Null, $mFont = Null, $iAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bDropdown = Null, $iLines = Null, $bMultiSel = Null, $aiDefaultSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[23]

	If Not IsObj($oListBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oListBox) <> $LOW_FORM_CON_TYPE_LIST_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $oLabelField, $iTxtDir, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $iMouseScroll, $bTabStop, $iTabOrder, $asList, $mFont, $iAlign, $iBackColor, $iBorder, $iBorderColor, $bDropdown, $iLines, $bMultiSel, $aiDefaultSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oListBox.Control.Name(), __LOWriter_FormConGetObj($oListBox.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oListBox.Control.WritingMode(), $oListBox.Control.Enabled(), _
				$oListBox.Control.EnableVisible(), $oListBox.Control.ReadOnly(), $oListBox.Control.Printable(), $oListBox.Control.MouseWheelBehavior(), $oListBox.Control.Tabstop(), _
				$oListBox.Control.TabIndex(), $oListBox.Control.StringItemList(), __LOWriter_FormConSetGetFontDesc($oListBox), $oListBox.Control.Align(), $oListBox.Control.BackgroundColor(), _
				$oListBox.Control.Border(), $oListBox.Control.BorderColor(), $oListBox.Control.Dropdown(), $oListBox.Control.LineCount(), $oListBox.Control.MultiSelection(), _
				$oListBox.Control.DefaultSelection(), $oListBox.Control.Tag(), $oListBox.Control.HelpText(), $oListBox.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oListBox.Control.Name = $sName
		$iError = ($oListBox.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oLabelField = Default) Then
		$oListBox.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oListBox.Control.LabelControl = $oLabelField.Control()
		$iError = ($oListBox.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oListBox.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oListBox.Control.WritingMode = $iTxtDir
		$iError = ($oListBox.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bEnabled = Default) Then
		$oListBox.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oListBox.Control.Enabled = $bEnabled
		$iError = ($oListBox.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bVisible = Default) Then
		$oListBox.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oListBox.Control.EnableVisible = $bVisible
		$iError = ($oListBox.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bReadOnly = Default) Then
		$oListBox.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oListBox.Control.ReadOnly = $bReadOnly
		$iError = ($oListBox.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bPrintable = Default) Then
		$oListBox.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oListBox.Control.Printable = $bPrintable
		$iError = ($oListBox.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iMouseScroll = Default) Then
		$oListBox.Control.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oListBox.Control.MouseWheelBehavior = $iMouseScroll
		$iError = ($oListBox.Control.MouseWheelBehavior = $iMouseScroll) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bTabStop = Default) Then
		$oListBox.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oListBox.Control.Tabstop = $bTabStop
		$iError = ($oListBox.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 512) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oListBox.Control.TabIndex = $iTabOrder
		$iError = ($oListBox.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($asList = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default StringItemList.

	ElseIf ($asList <> Null) Then
		If Not IsArray($asList) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		For $i = 0 To UBound($asList) - 1
			If Not IsString($asList[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, $i)

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next
		$oListBox.Control.StringItemList = $asList
		$iError = (UBound($oListBox.Control.StringItemList()) = UBound($asList)) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 2048) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		__LOWriter_FormConSetGetFontDesc($oListBox, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($iAlign = Default) Then
		$oListBox.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oListBox.Control.Align = $iAlign
		$iError = ($oListBox.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iBackColor = Default) Then
		$oListBox.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oListBox.Control.BackgroundColor = $iBackColor
		$iError = ($oListBox.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iBorder = Default) Then
		$oListBox.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oListBox.Control.Border = $iBorder
		$iError = ($oListBox.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iBorderColor = Default) Then
		$oListBox.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oListBox.Control.BorderColor = $iBorderColor
		$iError = ($oListBox.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bDropdown = Default) Then
		$oListBox.Control.setPropertyToDefault("Dropdown")

	ElseIf ($bDropdown <> Null) Then
		If Not IsBool($bDropdown) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oListBox.Control.Dropdown = $bDropdown
		$iError = ($oListBox.Control.Dropdown() = $bDropdown) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($iLines = Default) Then
		$oListBox.Control.setPropertyToDefault("LineCount")

	ElseIf ($iLines <> Null) Then
		If Not __LO_IntIsBetween($iLines, -2147483648, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oListBox.Control.LineCount = $iLines
		$iError = ($oListBox.Control.LineCount = $iLines) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($bMultiSel = Default) Then
		$oListBox.Control.setPropertyToDefault("MultiSelection")

	ElseIf ($bMultiSel <> Null) Then
		If Not IsBool($bMultiSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oListBox.Control.MultiSelection = $bMultiSel
		$iError = ($oListBox.Control.MultiSelection() = $bMultiSel) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($aiDefaultSel = Default) Then
		$iError = BitOR($iError, 524288) ; Can't Default Name.

	ElseIf ($aiDefaultSel <> Null) Then
		If Not IsArray($aiDefaultSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		For $i = 0 To UBound($aiDefaultSel) - 1
			If Not IsInt($aiDefaultSel[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, $i)
			If ($aiDefaultSel[$i] >= $oListBox.Control.ItemCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, $i)

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next
		$oListBox.Control.DefaultSelection = $aiDefaultSel
		$iError = (UBound($oListBox.Control.DefaultSelection()) = UBound($aiDefaultSel)) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 1048576) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 27, 0)

		$oListBox.Control.Tag = $sAddInfo
		$iError = ($oListBox.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($sHelpText = Default) Then
		$oListBox.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 28, 0)

		$oListBox.Control.HelpText = $sHelpText
		$iError = ($oListBox.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($sHelpURL = Default) Then
		$oListBox.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 29, 0)

		$oListBox.Control.HelpURL = $sHelpURL
		$iError = ($oListBox.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConListBoxGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConListBoxGetCount
; Description ...: Retrieve a count of values contained in a List Box.
; Syntax ........: _LOWriter_FormConListBoxGetCount(ByRef $oListBox)
; Parameters ....: $oListBox            - [in/out] an object. A List Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oListBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oListBox not a List Box Control.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Item Count.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning the count of List Box entries.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormConListBoxGeneral, _LOWriter_FormConListBoxData, _LOWriter_FormConListBoxSelection
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConListBoxGetCount(ByRef $oListBox)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oListBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oListBox) <> $LOW_FORM_CON_TYPE_LIST_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oListBox.Control.ItemCount()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOWriter_FormConListBoxGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConListBoxSelection
; Description ...: Set or Retrieve the current List Box selection.
; Syntax ........: _LOWriter_FormConListBoxSelection(ByRef $oListBox[, $aiSelection = Null[, $bReturnValue = False]])
; Parameters ....: $oListBox            - [in/out] an object. A List Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $aiSelection         - [optional] an array of integers. Default is Null. A single dimension array of selection values. See remarks.
;                  $bReturnValue        - [optional] a boolean value. Default is False. If True, when retrieving the the current selection(s), the current selected VALUE, instead of the position is returned.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oListBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oListBox not a List Box Control.
;                  @Error 1 @Extended 3 Return 0 = $bReturnValue not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $aiSelection not an array.
;                  @Error 1 @Extended 5 Return ? = Array called in $aiSelection contains an element with a non-Integer value. Returning problem element position.
;                  @Error 1 @Extended 6 Return ? = Array called in $aiSelection contains an element with an Integer value less than 0 or greater than number of List Box entries. Returning problem element position.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current selection.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $aiSelection
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current selection(s) of the List Box. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The array called for $aiSelection should be a single dimension array, with one Integer value, corresponding to the position in the List box value array, per array element, to indicate which value(s) is/are selected.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current Selection(s) of the List Box. If $bReturnValue is False, the return will be a single dimension array with each element containing an Integer indicating which List Box value is selected, else if $bReturnValue is True, a single dimension array will be returned, with each element containing a selected value.
;                  Call $aiSelection with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConListBoxGeneral, _LOWriter_FormConListBoxData, _LOWriter_FormConListBoxGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConListBoxSelection(ByRef $oListBox, $aiSelection = Null, $bReturnValue = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiCurSel[0], $avCurSel[0]

	If Not IsObj($oListBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oListBox) <> $LOW_FORM_CON_TYPE_LIST_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not IsBool($bReturnValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($aiSelection) Then
		If $bReturnValue Then
			$avCurSel = $oListBox.Control.SelectedValues()
			If Not IsArray($avCurSel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			Return SetError($__LO_STATUS_SUCCESS, 1, $avCurSel)

		Else
			$aiCurSel = $oListBox.Control.SelectedItems()
			If Not IsArray($aiCurSel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			Return SetError($__LO_STATUS_SUCCESS, 2, $aiCurSel)
		EndIf
	EndIf

	If ($aiSelection = Default) Then
		$oListBox.Control.setPropertyToDefault("SelectedItems")

	Else
		If Not IsArray($aiSelection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		For $i = 0 To UBound($aiSelection) - 1
			If Not IsInt($aiSelection[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
			If Not __LO_IntIsBetween($aiSelection[$i], 0, $oListBox.Control.ItemCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		$oListBox.Control.SelectedItems = $aiSelection
		$iError = (UBound($oListBox.Control.SelectedItems()) = UBound($aiSelection)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConListBoxSelection

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConNavBarGeneral
; Description ...: Set or Retrieve general Navigation Bar properties.
; Syntax ........: _LOWriter_FormConNavBarGeneral(ByRef $oNavBar[, $sName = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bVisible = Null[, $bTabStop = Null[, $iTabOrder = Null[, $iDelay = Null[, $mFont = Null[, $iBackColor = Null[, $iBorder = Null[, $bSmallIcon = Null[, $bShowPos = Null[, $bShowNav = Null[, $bShowActing = Null[, $bShowFiltering = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]])
; Parameters ....: $oNavBar             - [in/out] an object. A Navigation Bar Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $bSmallIcon          - [optional] a boolean value. Default is Null. If True, small Icon sizing is used.
;                  $bShowPos            - [optional] a boolean value. Default is Null. If True, positioning items are shown.
;                  $bShowNav            - [optional] a boolean value. Default is Null. If True, navigation items are shown.
;                  $bShowActing         - [optional] a boolean value. Default is Null. If True, action items are shown.
;                  $bShowFiltering      - [optional] a boolean value. Default is Null. If True, filtering and sorting items are shown.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oNavBar not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oNavBar not a Navigation Bar Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 9 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 10 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 11 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 12 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $bSmallIcon not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $bShowPos not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $bShowNav not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $bShowActing not a Boolean.
;                  @Error 1 @Extended 17 Return 0 = $bShowFiltering not a Boolean.
;                  @Error 1 @Extended 18 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 19 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 20 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $iTxtDir
;                  |                               4 = Error setting $bEnabled
;                  |                               8 = Error setting $bVisible
;                  |                               16 = Error setting $bTabStop
;                  |                               32 = Error setting $iTabOrder
;                  |                               64 = Error setting $iDelay
;                  |                               128 = Error setting $mFont
;                  |                               256 = Error setting $iBackColor
;                  |                               512 = Error setting $iBorder
;                  |                               1024 = Error setting $bSmallIcon
;                  |                               2048 = Error setting $bShowPos
;                  |                               4096 = Error setting $bShowNav
;                  |                               8192 = Error setting $bShowActing
;                  |                               16384 = Error setting $bShowFiltering
;                  |                               32768 = Error setting $sAddInfo
;                  |                               65536 = Error setting $sHelpText
;                  |                               131072 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 18 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConNavBarGeneral(ByRef $oNavBar, $sName = Null, $iTxtDir = Null, $bEnabled = Null, $bVisible = Null, $bTabStop = Null, $iTabOrder = Null, $iDelay = Null, $mFont = Null, $iBackColor = Null, $iBorder = Null, $bSmallIcon = Null, $bShowPos = Null, $bShowNav = Null, $bShowActing = Null, $bShowFiltering = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[18]
	Local Const $__LOW_FORM_CONTROL_ICON_SMALL = 0, $__LOW_FORM_CONTROL_ICON_LARGE = 1

	If Not IsObj($oNavBar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oNavBar) <> $LOW_FORM_CON_TYPE_NAV_BAR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $iTxtDir, $bEnabled, $bVisible, $bTabStop, $iTabOrder, $iDelay, $mFont, $iBackColor, $iBorder, $bSmallIcon, $bShowPos, $bShowNav, $bShowActing, $bShowFiltering, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oNavBar.Control.Name(), $oNavBar.Control.WritingMode(), $oNavBar.Control.Enabled(), $oNavBar.Control.EnableVisible(), _
				$oNavBar.Control.Tabstop(), $oNavBar.Control.TabIndex(), $oNavBar.Control.RepeatDelay(), __LOWriter_FormConSetGetFontDesc($oNavBar), $oNavBar.Control.BackgroundColor(), _
				$oNavBar.Control.Border(), (($oNavBar.Control.IconSize() = $__LOW_FORM_CONTROL_ICON_SMALL) ? (True) : (False)), _ ; Icon size.
				$oNavBar.Control.ShowPosition(), $oNavBar.Control.ShowNavigation(), $oNavBar.Control.ShowRecordActions(), $oNavBar.Control.ShowFilterSort(), _
				$oNavBar.Control.Tag(), $oNavBar.Control.HelpText(), $oNavBar.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oNavBar.Control.Name = $sName
		$iError = ($oNavBar.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTxtDir = Default) Then
		$oNavBar.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oNavBar.Control.WritingMode = $iTxtDir
		$iError = ($oNavBar.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bEnabled = Default) Then
		$oNavBar.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oNavBar.Control.Enabled = $bEnabled
		$iError = ($oNavBar.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bVisible = Default) Then
		$oNavBar.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oNavBar.Control.EnableVisible = $bVisible
		$iError = ($oNavBar.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bTabStop = Default) Then
		$oNavBar.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oNavBar.Control.Tabstop = $bTabStop
		$iError = ($oNavBar.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 32) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oNavBar.Control.TabIndex = $iTabOrder
		$iError = ($oNavBar.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iDelay = Default) Then
		$oNavBar.Control.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oNavBar.Control.RepeatDelay = $iDelay
		$iError = ($oNavBar.Control.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 128) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		__LOWriter_FormConSetGetFontDesc($oNavBar, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iBackColor = Default) Then
		$oNavBar.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oNavBar.Control.BackgroundColor = $iBackColor
		$iError = ($oNavBar.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iBorder = Default) Then
		$oNavBar.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oNavBar.Control.Border = $iBorder
		$iError = ($oNavBar.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($bSmallIcon = Default) Then
		$oNavBar.Control.setPropertyToDefault("IconSize")

	ElseIf ($bSmallIcon <> Null) Then
		If Not IsBool($bSmallIcon) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oNavBar.Control.IconSize = (($bSmallIcon) ? ($__LOW_FORM_CONTROL_ICON_SMALL) : ($__LOW_FORM_CONTROL_ICON_LARGE))
		$iError = ($oNavBar.Control.IconSize() = (($bSmallIcon) ? ($__LOW_FORM_CONTROL_ICON_SMALL) : ($__LOW_FORM_CONTROL_ICON_LARGE))) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($bShowPos = Default) Then
		$oNavBar.Control.setPropertyToDefault("ShowPosition")

	ElseIf ($bShowPos <> Null) Then
		If Not IsBool($bShowPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oNavBar.Control.ShowPosition = $bShowPos
		$iError = ($oNavBar.Control.ShowPosition() = $bShowPos) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($bShowNav = Default) Then
		$oNavBar.Control.setPropertyToDefault("ShowNavigation")

	ElseIf ($bShowNav <> Null) Then
		If Not IsBool($bShowNav) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oNavBar.Control.ShowNavigation = $bShowNav
		$iError = ($oNavBar.Control.ShowNavigation() = $bShowNav) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($bShowActing = Default) Then
		$oNavBar.Control.setPropertyToDefault("ShowRecordActions")

	ElseIf ($bShowActing <> Null) Then
		If Not IsBool($bShowActing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oNavBar.Control.ShowRecordActions = $bShowActing
		$iError = ($oNavBar.Control.ShowRecordActions() = $bShowActing) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($bShowFiltering = Default) Then
		$oNavBar.Control.setPropertyToDefault("ShowFilterSort")

	ElseIf ($bShowFiltering <> Null) Then
		If Not IsBool($bShowFiltering) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oNavBar.Control.ShowFilterSort = $bShowFiltering
		$iError = ($oNavBar.Control.ShowFilterSort() = $bShowFiltering) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 32768) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oNavBar.Control.Tag = $sAddInfo
		$iError = ($oNavBar.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($sHelpText = Default) Then
		$oNavBar.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oNavBar.Control.HelpText = $sHelpText
		$iError = ($oNavBar.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($sHelpURL = Default) Then
		$oNavBar.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oNavBar.Control.HelpURL = $sHelpURL
		$iError = ($oNavBar.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConNavBarGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConNumericFieldData
; Description ...: Set or Retrieve Numeric Field Data Properties.
; Syntax ........: _LOWriter_FormConNumericFieldData(ByRef $oNumericField[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oNumericField       - [in/out] an object. A Numeric Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oNumericField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oNumericField not a Numeric Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConNumericFieldValue, _LOWriter_FormConNumericFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConNumericFieldData(ByRef $oNumericField, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oNumericField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oNumericField) <> $LOW_FORM_CON_TYPE_NUMERIC_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oNumericField.Control.DataField(), $oNumericField.Control.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oNumericField.Control.DataField = $sDataField
		$iError = ($oNumericField.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oNumericField.Control.InputRequired = $bInputRequired
		$iError = ($oNumericField.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConNumericFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConNumericFieldGeneral
; Description ...: Set or Retrieve general Numeric Field properties.
; Syntax ........: _LOWriter_FormConNumericFieldGeneral(ByRef $oNumericField[, $sName = Null[, $oLabelField = Null[, $iTxtDir = Null[, $bStrict = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $iMouseScroll = Null[, $bTabStop = Null[, $iTabOrder = Null[, $nMin = Null[, $nMax = Null[, $iIncr = Null[, $nDefault = Null[, $iDecimal = Null[, $bThousandsSep = Null[, $bSpin = Null[, $bRepeat = Null[, $iDelay = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oNumericField       - [in/out] an object. A Numeric Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bStrict             - [optional] a boolean value. Default is Null. If True, strict formatting is enabled.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $nMin                - [optional] a general number value. Default is Null. The minimum value the control can be set to.
;                  $nMax                - [optional] a general number value. Default is Null. The maximum value the control can be set to.
;                  $iIncr               - [optional] an integer value. Default is Null. The amount to Increase or Decrease the value by.
;                  $nDefault            - [optional] a general number value. Default is Null. The default value the control will be set to.
;                  $iDecimal            - [optional] an integer value (0-20). Default is Null. The amount of decimal accuracy.
;                  $bThousandsSep       - [optional] a boolean value. Default is Null. If True, a thousands separator will be added.
;                  $bSpin               - [optional] a boolean value. Default is Null. If True, the field will act as a spin button.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oNumericField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oNumericField not a Numeric Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 5 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 6 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bStrict not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 15 Return 0 = $nMin not a Number.
;                  @Error 1 @Extended 16 Return 0 = $nMax not a Number.
;                  @Error 1 @Extended 17 Return 0 = $iIncr not an Integer.
;                  @Error 1 @Extended 18 Return 0 = $nDefault not a Number.
;                  @Error 1 @Extended 19 Return 0 = $iDecimal not an Integer, less than 0 or greater than 20.
;                  @Error 1 @Extended 20 Return 0 = $bThousandsSep not a Boolean.
;                  @Error 1 @Extended 21 Return 0 = $bSpin not a Boolean.
;                  @Error 1 @Extended 22 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 23 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 24 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 25 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 26 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 27 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 28 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 29 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 30 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 31 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 32 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 33 Return 0 = $sHelpURL not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.DateTime" Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.util.Time" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current Minimum Time.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve current Maximum Time.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify parent document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $oLabelField
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bStrict
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bVisible
;                  |                               64 = Error setting $bReadOnly
;                  |                               128 = Error setting $bPrintable
;                  |                               256 = Error setting $iMouseScroll
;                  |                               512 = Error setting $bTabStop
;                  |                               1024 = Error setting $iTabOrder
;                  |                               2048 = Error setting $nMin
;                  |                               4096 = Error setting $nMax
;                  |                               8192 = Error setting $iIncr
;                  |                               16384 = Error setting $nDefault
;                  |                               32768 = Error setting $iDecimal
;                  |                               65536 = Error setting $bThousandsSep
;                  |                               131072 = Error setting $bSpin
;                  |                               262144 = Error setting $bRepeat
;                  |                               524288 = Error setting $iDelay
;                  |                               1048576 = Error setting $mFont
;                  |                               2097152 = Error setting $iAlign
;                  |                               4194304 = Error setting $iVertAlign
;                  |                               8388608 = Error setting $iBackColor
;                  |                               16777216 = Error setting $iBorder
;                  |                               33554432 = Error setting $iBorderColor
;                  |                               67108864 = Error setting $bHideSel
;                  |                               134217728 = Error setting $sAddInfo
;                  |                               268435456 = Error setting $sHelpText
;                  |                               536870912 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 30 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormConNumericFieldValue, _LOWriter_FormConNumericFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConNumericFieldGeneral(ByRef $oNumericField, $sName = Null, $oLabelField = Null, $iTxtDir = Null, $bStrict = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $iMouseScroll = Null, $bTabStop = Null, $iTabOrder = Null, $nMin = Null, $nMax = Null, $iIncr = Null, $nDefault = Null, $iDecimal = Null, $bThousandsSep = Null, $bSpin = Null, $bRepeat = Null, $iDelay = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[30]

	If Not IsObj($oNumericField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oNumericField) <> $LOW_FORM_CON_TYPE_NUMERIC_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $oLabelField, $iTxtDir, $bStrict, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $iMouseScroll, $bTabStop, $iTabOrder, $nMin, $nMax, $iIncr, $nDefault, $iDecimal, $bThousandsSep, $bSpin, $bRepeat, $iDelay, $mFont, $iAlign, $iVertAlign, $iBackColor, $iBorder, $iBorderColor, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oNumericField.Control.Name(), __LOWriter_FormConGetObj($oNumericField.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oNumericField.Control.WritingMode(), $oNumericField.Control.StrictFormat(), _
				$oNumericField.Control.Enabled(), $oNumericField.Control.EnableVisible(), $oNumericField.Control.ReadOnly(), $oNumericField.Control.Printable(), _
				$oNumericField.Control.MouseWheelBehavior(), $oNumericField.Control.Tabstop(), $oNumericField.Control.TabIndex(), $oNumericField.Control.ValueMin(), _
				$oNumericField.Control.ValueMax(), $oNumericField.Control.ValueStep(), $oNumericField.Control.DefaultValue(), $oNumericField.Control.DecimalAccuracy(), _
				$oNumericField.Control.ShowThousandsSeparator(), $oNumericField.Control.Spin(), $oNumericField.Control.Repeat(), $oNumericField.Control.RepeatDelay(), _
				__LOWriter_FormConSetGetFontDesc($oNumericField), $oNumericField.Control.Align(), $oNumericField.Control.VerticalAlign(), $oNumericField.Control.BackgroundColor(), _
				$oNumericField.Control.Border(), $oNumericField.Control.BorderColor(), $oNumericField.Control.HideInactiveSelection(), $oNumericField.Control.Tag(), _
				$oNumericField.Control.HelpText(), $oNumericField.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oNumericField.Control.Name = $sName
		$iError = ($oNumericField.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oLabelField = Default) Then
		$oNumericField.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oNumericField.Control.LabelControl = $oLabelField.Control()
		$iError = ($oNumericField.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oNumericField.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oNumericField.Control.WritingMode = $iTxtDir
		$iError = ($oNumericField.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bStrict = Default) Then
		$oNumericField.Control.setPropertyToDefault("StrictFormat")

	ElseIf ($bStrict <> Null) Then
		If Not IsBool($bStrict) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oNumericField.Control.StrictFormat = $bStrict
		$iError = ($oNumericField.Control.StrictFormat() = $bStrict) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oNumericField.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oNumericField.Control.Enabled = $bEnabled
		$iError = ($oNumericField.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bVisible = Default) Then
		$oNumericField.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oNumericField.Control.EnableVisible = $bVisible
		$iError = ($oNumericField.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bReadOnly = Default) Then
		$oNumericField.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oNumericField.Control.ReadOnly = $bReadOnly
		$iError = ($oNumericField.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bPrintable = Default) Then
		$oNumericField.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oNumericField.Control.Printable = $bPrintable
		$iError = ($oNumericField.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iMouseScroll = Default) Then
		$oNumericField.Control.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oNumericField.Control.MouseWheelBehavior = $iMouseScroll
		$iError = ($oNumericField.Control.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bTabStop = Default) Then
		$oNumericField.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oNumericField.Control.Tabstop = $bTabStop
		$iError = ($oNumericField.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oNumericField.Control.TabIndex = $iTabOrder
		$iError = ($oNumericField.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($nMin = Default) Then
		$oNumericField.Control.setPropertyToDefault("ValueMin")

	ElseIf ($nMin <> Null) Then
		If Not IsNumber($nMin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oNumericField.Control.ValueMin = $nMin
		$iError = ($oNumericField.Control.ValueMin() = $nMin) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($nMax = Default) Then
		$oNumericField.Control.setPropertyToDefault("ValueMax")

	ElseIf ($nMax <> Null) Then
		If Not IsNumber($nMax) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oNumericField.Control.ValueMax = $nMax
		$iError = ($oNumericField.Control.ValueMax() = $nMax) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iIncr = Default) Then
		$oNumericField.Control.setPropertyToDefault("ValueStep")

	ElseIf ($iIncr <> Null) Then
		If Not IsInt($iIncr) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oNumericField.Control.ValueStep = $iIncr
		$iError = ($oNumericField.Control.ValueStep() = $iIncr) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($nDefault = Default) Then
		$oNumericField.Control.setPropertyToDefault("DefaultValue")

	ElseIf ($nDefault <> Null) Then
		If Not IsNumber($nDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oNumericField.Control.DefaultValue = $nDefault
		$iError = ($oNumericField.Control.DefaultValue() = $nDefault) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iDecimal = Default) Then
		$oNumericField.Control.setPropertyToDefault("DecimalAccuracy")

	ElseIf ($iDecimal <> Null) Then
		If Not __LO_IntIsBetween($iDecimal, 0, 20) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oNumericField.Control.DecimalAccuracy = $iDecimal
		$iError = ($oNumericField.Control.DecimalAccuracy() = $iDecimal) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bThousandsSep = Default) Then
		$oNumericField.Control.setPropertyToDefault("ShowThousandsSeparator")

	ElseIf ($bThousandsSep <> Null) Then
		If Not IsBool($bThousandsSep) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oNumericField.Control.ShowThousandsSeparator = $bThousandsSep
		$iError = ($oNumericField.Control.ShowThousandsSeparator() = $bThousandsSep) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($bSpin = Default) Then
		$oNumericField.Control.setPropertyToDefault("Spin")

	ElseIf ($bSpin <> Null) Then
		If Not IsBool($bSpin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oNumericField.Control.Spin = $bSpin
		$iError = ($oNumericField.Control.Spin() = $bSpin) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($bRepeat = Default) Then
		$oNumericField.Control.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oNumericField.Control.Repeat = $bRepeat
		$iError = ($oNumericField.Control.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($iDelay = Default) Then
		$oNumericField.Control.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oNumericField.Control.RepeatDelay = $iDelay
		$iError = ($oNumericField.Control.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 1048576) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		__LOWriter_FormConSetGetFontDesc($oNumericField, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($iAlign = Default) Then
		$oNumericField.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oNumericField.Control.Align = $iAlign
		$iError = ($oNumericField.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($iVertAlign = Default) Then
		$oNumericField.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, 0)

		$oNumericField.Control.VerticalAlign = $iVertAlign
		$iError = ($oNumericField.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	If ($iBackColor = Default) Then
		$oNumericField.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 27, 0)

		$oNumericField.Control.BackgroundColor = $iBackColor
		$iError = ($oNumericField.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 8388608))
	EndIf

	If ($iBorder = Default) Then
		$oNumericField.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 28, 0)

		$oNumericField.Control.Border = $iBorder
		$iError = ($oNumericField.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 16777216))
	EndIf

	If ($iBorderColor = Default) Then
		$oNumericField.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 29, 0)

		$oNumericField.Control.BorderColor = $iBorderColor
		$iError = ($oNumericField.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 33554432))
	EndIf

	If ($bHideSel = Default) Then
		$oNumericField.Control.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 30, 0)

		$oNumericField.Control.HideInactiveSelection = $bHideSel
		$iError = ($oNumericField.Control.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 67108864))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 134217728) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 31, 0)

		$oNumericField.Control.Tag = $sAddInfo
		$iError = ($oNumericField.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 134217728))
	EndIf

	If ($sHelpText = Default) Then
		$oNumericField.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 32, 0)

		$oNumericField.Control.HelpText = $sHelpText
		$iError = ($oNumericField.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 268435456))
	EndIf

	If ($sHelpURL = Default) Then
		$oNumericField.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 33, 0)

		$oNumericField.Control.HelpURL = $sHelpURL
		$iError = ($oNumericField.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 536870912))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConNumericFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConNumericFieldValue
; Description ...: Set or retrieve the current Numeric field value.
; Syntax ........: _LOWriter_FormConNumericFieldValue(ByRef $oNumericField[, $nValue = Null])
; Parameters ....: $oNumericField       - [in/out] an object. A Numeric Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $nValue              - [optional] a general number value. Default is Null. The value to set the field to.
; Return values .: Success: 1 or Number
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oNumericField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oNumericField not a Numeric Field Control.
;                  @Error 1 @Extended 3 Return 0 = $nValue not a Number.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $nValue
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Number = Success. All optional parameters were called with Null, returning current value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current value. Return will be Null if a value hasn't been set.
;                  Call $nValue with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConNumericFieldGeneral, _LOWriter_FormConNumericFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConNumericFieldValue(ByRef $oNumericField, $nValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $nCurVal

	If Not IsObj($oNumericField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oNumericField) <> $LOW_FORM_CON_TYPE_NUMERIC_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($nValue) Then
		$nCurVal = $oNumericField.Control.Value() ; Value is Null when not set.

		Return SetError($__LO_STATUS_SUCCESS, 1, $nCurVal)
	EndIf

	If ($nValue = Default) Then
		$oNumericField.Control.setPropertyToDefault("Value")

	Else
		If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oNumericField.Control.Value = $nValue
		$iError = ($oNumericField.Control.Value() = $nValue) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConNumericFieldValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConOptionButtonData
; Description ...: Set or Retrieve Option Button Data Properties.
; Syntax ........: _LOWriter_FormConOptionButtonData(ByRef $oOptionButton[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oOptionButton       - [in/out] an object. A Option Button Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oOptionButton not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oOptionButton not a Option Button Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
;                  Reference Values are not included here as they are applicable to Calc only, as far as I can ascertain.
; Related .......: _LOWriter_FormConOptionButtonState, _LOWriter_FormConOptionButtonGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConOptionButtonData(ByRef $oOptionButton, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oOptionButton) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oOptionButton) <> $LOW_FORM_CON_TYPE_OPTION_BUTTON) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oOptionButton.Control.DataField(), $oOptionButton.Control.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oOptionButton.Control.DataField = $sDataField
		$iError = ($oOptionButton.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oOptionButton.Control.InputRequired = $bInputRequired
		$iError = ($oOptionButton.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConOptionButtonData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConOptionButtonGeneral
; Description ...: Set or Retrieve general Option button properties.
; Syntax ........: _LOWriter_FormConOptionButtonGeneral(ByRef $oOptionButton[, $sName = Null[, $sLabel = Null[, $oLabelField = Null[, $iTxtDir = Null[, $sGroupName = Null[, $bEnabled = Null[, $bVisible = Null[, $bPrintable = Null[, $bTabStop = Null[, $iTabOrder = Null[, $iDefaultState = Null[, $mFont = Null[, $iStyle = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $bWordBreak = Null[, $sGraphics = Null[, $iGraphicAlign = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oOptionButton       - [in/out] an object. A Option button Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $oLabelField         - [optional] an object. Default is Null. A Group Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sGroupName          - [optional] a string value. Default is Null. The Group name the control is in.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $iDefaultState       - [optional] an integer value (0-1). Default is Null. The Default state of the Option button. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iStyle              - [optional] an integer value (1-2). Default is Null. The display style of the Option button. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bWordBreak          - [optional] a boolean value. Default is Null. If True, line breaks are allowed to be used.
;                  $sGraphics           - [optional] a string value. Default is Null. The path to an Image file.
;                  $iGraphicAlign       - [optional] an integer value (0-12). Default is Null. The Alignment of the Image. See Constants $LOW_FORM_CON_IMG_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oOptionButton not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oOptionButton not a Option Button Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 6 Return 0 = Object called in $oLabelField not a Group Box Control.
;                  @Error 1 @Extended 7 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $sGroupName not a String.
;                  @Error 1 @Extended 9 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 13 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 14 Return 0 = $iDefaultState not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 15 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 16 Return 0 = $iStyle not an Integer, less than 1 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 17 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 18 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 19 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 20 Return 0 = $bWordBreak not a Boolean.
;                  @Error 1 @Extended 21 Return 0 = $sGraphics not a String.
;                  @Error 1 @Extended 22 Return 0 = $iGraphicAlign not an Integer, less than 0 or greater than 12. See Constants $LOW_FORM_CON_IMG_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 23 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 24 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 25 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $oLabelField
;                  |                               8 = Error setting $iTxtDir
;                  |                               16 = Error setting $sGroupName
;                  |                               32 = Error setting $bEnabled
;                  |                               64 = Error setting $bVisible
;                  |                               128 = Error setting $bPrintable
;                  |                               256 = Error setting $bTabStop
;                  |                               512 = Error setting $iTabOrder
;                  |                               1024 = Error setting $iDefaultState
;                  |                               2048 = Error setting $mFont
;                  |                               4096 = Error setting $iStyle
;                  |                               8192 = Error setting $iAlign
;                  |                               16384 = Error setting $iVertAlign
;                  |                               32768 = Error setting $iBackColor
;                  |                               65536 = Error setting $bWordBreak
;                  |                               131072 = Error setting $sGraphics
;                  |                               262144 = Error setting $iGraphicAlign
;                  |                               524288 = Error setting $sAddInfo
;                  |                               1048576 = Error setting $sHelpText
;                  |                               2097152 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 22 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $sGraphics is called with an invalid Graphic URL, graphic is set to Null. The Return for $sGraphics is an Image Object.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $iDefaultState, $mFont, $sAddInfo.
; Related .......: _LOWriter_FormConOptionButtonState, _LOWriter_FormConOptionButtonData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConOptionButtonGeneral(ByRef $oOptionButton, $sName = Null, $sLabel = Null, $oLabelField = Null, $iTxtDir = Null, $sGroupName = Null, $bEnabled = Null, $bVisible = Null, $bPrintable = Null, $bTabStop = Null, $iTabOrder = Null, $iDefaultState = Null, $mFont = Null, $iStyle = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $bWordBreak = Null, $sGraphics = Null, $iGraphicAlign = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[22]

	If Not IsObj($oOptionButton) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oOptionButton) <> $LOW_FORM_CON_TYPE_OPTION_BUTTON) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $oLabelField, $iTxtDir, $sGroupName, $bEnabled, $bVisible, $bPrintable, $bTabStop, $iTabOrder, $iDefaultState, $mFont, $iStyle, $iAlign, $iVertAlign, $iBackColor, $bWordBreak, $sGraphics, $iGraphicAlign, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oOptionButton.Control.Name(), $oOptionButton.Control.Label(), __LOWriter_FormConGetObj($oOptionButton.Control.LabelControl(), $LOW_FORM_CON_TYPE_GROUP_BOX), $oOptionButton.Control.WritingMode(), _
				$oOptionButton.Control.GroupName(), $oOptionButton.Control.Enabled(), $oOptionButton.Control.EnableVisible(), $oOptionButton.Control.Printable(), _
				$oOptionButton.Control.Tabstop(), $oOptionButton.Control.TabIndex(), $oOptionButton.Control.DefaultState(), __LOWriter_FormConSetGetFontDesc($oOptionButton), _
				$oOptionButton.Control.VisualEffect(), $oOptionButton.Control.Align(), $oOptionButton.Control.VerticalAlign(), $oOptionButton.Control.BackgroundColor(), _
				$oOptionButton.Control.MultiLine(), $oOptionButton.Control.Graphic(), $oOptionButton.Control.ImagePosition(), $oOptionButton.Control.Tag(), _
				$oOptionButton.Control.HelpText(), $oOptionButton.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oOptionButton.Control.Name = $sName
		$iError = ($oOptionButton.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oOptionButton.Control.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oOptionButton.Control.Label = $sLabel
		$iError = ($oOptionButton.Control.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oLabelField = Default) Then
		$oOptionButton.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_GROUP_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oOptionButton.Control.LabelControl = $oLabelField.Control()
		$iError = ($oOptionButton.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iTxtDir = Default) Then
		$oOptionButton.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oOptionButton.Control.WritingMode = $iTxtDir
		$iError = ($oOptionButton.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sGroupName = Default) Then
		$oOptionButton.Control.setPropertyToDefault("GroupName")

	ElseIf ($sGroupName <> Null) Then
		If Not IsString($sGroupName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oOptionButton.Control.GroupName = $sGroupName
		$iError = ($oOptionButton.Control.GroupName() = $sGroupName) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bEnabled = Default) Then
		$oOptionButton.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oOptionButton.Control.Enabled = $bEnabled
		$iError = ($oOptionButton.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bVisible = Default) Then
		$oOptionButton.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oOptionButton.Control.EnableVisible = $bVisible
		$iError = ($oOptionButton.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bPrintable = Default) Then
		$oOptionButton.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oOptionButton.Control.Printable = $bPrintable
		$iError = ($oOptionButton.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bTabStop = Default) Then
		$oOptionButton.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oOptionButton.Control.Tabstop = $bTabStop
		$iError = ($oOptionButton.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 512) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oOptionButton.Control.TabIndex = $iTabOrder
		$iError = ($oOptionButton.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iDefaultState = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default DefaultState.

	ElseIf ($iDefaultState <> Null) Then
		If Not __LO_IntIsBetween($iDefaultState, $LOW_FORM_CON_CHKBX_STATE_NOT_SELECTED, $LOW_FORM_CON_CHKBX_STATE_SELECTED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oOptionButton.Control.DefaultState = $iDefaultState
		$iError = ($oOptionButton.Control.DefaultState() = $iDefaultState) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 2048) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		__LOWriter_FormConSetGetFontDesc($oOptionButton, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($iStyle = Default) Then
		$oOptionButton.Control.setPropertyToDefault("VisualEffect")

	ElseIf ($iStyle <> Null) Then
		If Not __LO_IntIsBetween($iStyle, $LOW_FORM_CON_BORDER_3D, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oOptionButton.Control.VisualEffect = $iStyle
		$iError = ($oOptionButton.Control.VisualEffect() = $iStyle) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iAlign = Default) Then
		$oOptionButton.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oOptionButton.Control.Align = $iAlign
		$iError = ($oOptionButton.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iVertAlign = Default) Then
		$oOptionButton.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oOptionButton.Control.VerticalAlign = $iVertAlign
		$iError = ($oOptionButton.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iBackColor = Default) Then
		$oOptionButton.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oOptionButton.Control.BackgroundColor = $iBackColor
		$iError = ($oOptionButton.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bWordBreak = Default) Then
		$oOptionButton.Control.setPropertyToDefault("MultiLine")

	ElseIf ($bWordBreak <> Null) Then
		If Not IsBool($bWordBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oOptionButton.Control.MultiLine = $bWordBreak
		$iError = ($oOptionButton.Control.MultiLine() = $bWordBreak) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($sGraphics = Default) Then
		$oOptionButton.Control.setPropertyToDefault("ImageURL")
		$oOptionButton.Control.setPropertyToDefault("Graphic")

	ElseIf ($sGraphics <> Null) Then
		If Not IsString($sGraphics) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oOptionButton.Control.ImageURL = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)
		$iError = ($oOptionButton.Control.ImageURL() = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($iGraphicAlign = Default) Then
		$oOptionButton.Control.setPropertyToDefault("ImagePosition")

	ElseIf ($iGraphicAlign <> Null) Then
		If Not __LO_IntIsBetween($iGraphicAlign, $LOW_FORM_CON_IMG_ALIGN_LEFT_TOP, $LOW_FORM_CON_IMG_ALIGN_CENTERED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oOptionButton.Control.ImagePosition = $iGraphicAlign
		$iError = ($oOptionButton.Control.ImagePosition() = $iGraphicAlign) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 524288) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oOptionButton.Control.Tag = $sAddInfo
		$iError = ($oOptionButton.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($sHelpText = Default) Then
		$oOptionButton.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oOptionButton.Control.HelpText = $sHelpText
		$iError = ($oOptionButton.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($sHelpURL = Default) Then
		$oOptionButton.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oOptionButton.Control.HelpURL = $sHelpURL
		$iError = ($oOptionButton.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConOptionButtonGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConOptionButtonState
; Description ...: Set or Retrieve the current Option button state.
; Syntax ........: _LOWriter_FormConOptionButtonState(ByRef $oOptionButton[, $iState = Null])
; Parameters ....: $oOptionButton       - [in/out] an object. A Option Button Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iState              - [optional] an integer value (0-1). Default is Null. The current state of the Option Button. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oOptionButton not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oOptionButton not a Option Button Control.
;                  @Error 1 @Extended 3 Return 0 = $iState not an Integer, less than 0 or greater than 1. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current control State.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iState
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current Option Button State as an Integer, matching one of the constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current Option Button state.
;                  Call $iState with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConOptionButtonGeneral, _LOWriter_FormConOptionButtonData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConOptionButtonState(ByRef $oOptionButton, $iState = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurState

	If Not IsObj($oOptionButton) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oOptionButton) <> $LOW_FORM_CON_TYPE_OPTION_BUTTON) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iState) Then
		$iCurState = $oOptionButton.Control.State()
		If Not IsInt($iCurState) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurState)
	EndIf

	If ($iState = Default) Then
		$oOptionButton.Control.setPropertyToDefault("State")

	Else
		If Not __LO_IntIsBetween($iState, $LOW_FORM_CON_CHKBX_STATE_NOT_SELECTED, $LOW_FORM_CON_CHKBX_STATE_SELECTED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oOptionButton.Control.State = $iState
		$iError = ($oOptionButton.Control.State() = $iState) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConOptionButtonState

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConPatternFieldData
; Description ...: Set or Retrieve Pattern Field Data Properties.
; Syntax ........: _LOWriter_FormConPatternFieldData(ByRef $oPatternField[, $sDataField = Null[, $bEmptyIsNull = Null[, $bInputRequired = Null[, $bFilter = Null]]]])
; Parameters ....: $oPatternField       - [in/out] an object. A Pattern Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bEmptyIsNull        - [optional] a boolean value. Default is Null. If True, an empty string will be treated as a Null value.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
;                  $bFilter             - [optional] a boolean value. Default is Null. If True, filter proposal is active.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPatternField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oPatternField not a Pattern Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bEmptyIsNull not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bInputRequired not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bFilter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bEmptyIsNull
;                  |                               4 = Error setting $bInputRequired
;                  |                               8 = Error setting $bFilter
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConPatternFieldValue, _LOWriter_FormConPatternFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConPatternFieldData(ByRef $oPatternField, $sDataField = Null, $bEmptyIsNull = Null, $bInputRequired = Null, $bFilter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[4]

	If Not IsObj($oPatternField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oPatternField) <> $LOW_FORM_CON_TYPE_PATTERN_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bEmptyIsNull, $bInputRequired, $bFilter) Then
		__LO_ArrayFill($avControl, $oPatternField.Control.DataField(), $oPatternField.Control.ConvertEmptyToNull(), $oPatternField.Control.InputRequired(), _
				$oPatternField.Control.UseFilterValueProposal())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPatternField.Control.DataField = $sDataField
		$iError = ($oPatternField.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bEmptyIsNull <> Null) Then
		If Not IsBool($bEmptyIsNull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPatternField.Control.ConvertEmptyToNull = $bEmptyIsNull
		$iError = ($oPatternField.Control.ConvertEmptyToNull() = $bEmptyIsNull) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPatternField.Control.InputRequired = $bInputRequired
		$iError = ($oPatternField.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bFilter <> Null) Then
		If Not IsBool($bFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPatternField.Control.UseFilterValueProposal = $bFilter
		$iError = ($oPatternField.Control.UseFilterValueProposal() = $bFilter) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConPatternFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConPatternFieldGeneral
; Description ...: Set or Retrieve general Pattern Field properties.
; Syntax ........: _LOWriter_FormConPatternFieldGeneral(ByRef $oPatternField[, $sName = Null[, $oLabelField = Null[, $iTxtDir = Null[, $iMaxLen = Null[, $sEditMask = Null[, $sLiteralMask = Null[, $bStrict = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $iMouseScroll = Null[, $bTabStop = Null[, $iTabOrder = Null[, $sDefaultTxt = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oPatternField       - [in/out] an object. A Pattern Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iMaxLen             - [optional] an integer value (-1-2147483647). Default is Null. The maximum text length that the Pattern field will accept.
;                  $sEditMask           - [optional] a string value. Default is Null. The edit mask of the field.
;                  $sLiteralMask        - [optional] a string value. Default is Null. The literal mask of the field.
;                  $bStrict             - [optional] a boolean value. Default is Null. If True, strict formatting is enabled.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $sDefaultTxt         - [optional] a string value. Default is Null. The default text to display in the field.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPatternField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oPatternField not a Pattern Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 5 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 6 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $iMaxLen not an Integer, less than -1 or greater than 2147483647.
;                  @Error 1 @Extended 8 Return 0 = $sEditMask
;                  @Error 1 @Extended 9 Return 0 = $sLiteralMask not a String.
;                  @Error 1 @Extended 10 Return 0 = $bStrict not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 13 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 16 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 17 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 18 Return 0 = $sDefaultTxt not a String.
;                  @Error 1 @Extended 19 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 20 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 21 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 22 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 23 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 24 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 25 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 26 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 27 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 28 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $oLabelField
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $iMaxLen
;                  |                               16 = Error setting $sEditMask
;                  |                               32 = Error setting $sLiteralMask
;                  |                               64 = Error setting $bStrict
;                  |                               128 = Error setting $bEnabled
;                  |                               256 = Error setting $bVisible
;                  |                               512 = Error setting $bReadOnly
;                  |                               1024 = Error setting $bPrintable
;                  |                               2048 = Error setting $iMouseScroll
;                  |                               4096 = Error setting $bTabStop
;                  |                               8192 = Error setting $iTabOrder
;                  |                               16384 = Error setting $sDefaultTxt
;                  |                               32768 = Error setting $mFont
;                  |                               65536 = Error setting $iAlign
;                  |                               131072 = Error setting $iVertAlign
;                  |                               262144 = Error setting $iBackColor
;                  |                               524288 = Error setting $iBorder
;                  |                               1048576 = Error setting $iBorderColor
;                  |                               2097152 = Error setting $bHideSel
;                  |                               4194304 = Error setting $sAddInfo
;                  |                               8388608 = Error setting $sHelpText
;                  |                               16777216 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 25 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $sDefaultTxt, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormConPatternFieldValue, _LOWriter_FormConPatternFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConPatternFieldGeneral(ByRef $oPatternField, $sName = Null, $oLabelField = Null, $iTxtDir = Null, $iMaxLen = Null, $sEditMask = Null, $sLiteralMask = Null, $bStrict = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $iMouseScroll = Null, $bTabStop = Null, $iTabOrder = Null, $sDefaultTxt = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[25]

	If Not IsObj($oPatternField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oPatternField) <> $LOW_FORM_CON_TYPE_PATTERN_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $oLabelField, $iTxtDir, $iMaxLen, $sEditMask, $sLiteralMask, $bStrict, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $iMouseScroll, $bTabStop, $iTabOrder, $sDefaultTxt, $mFont, $iAlign, $iVertAlign, $iBackColor, $iBorder, $iBorderColor, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oPatternField.Control.Name(), __LOWriter_FormConGetObj($oPatternField.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oPatternField.Control.WritingMode(), _
				$oPatternField.Control.MaxTextLen(), $oPatternField.Control.EditMask(), $oPatternField.Control.LiteralMask(), $oPatternField.Control.StrictFormat(), _
				$oPatternField.Control.Enabled(), $oPatternField.Control.EnableVisible(), $oPatternField.Control.ReadOnly(), $oPatternField.Control.Printable(), _
				$oPatternField.Control.MouseWheelBehavior(), $oPatternField.Control.Tabstop(), $oPatternField.Control.TabIndex(), $oPatternField.Control.DefaultText(), _
				__LOWriter_FormConSetGetFontDesc($oPatternField), $oPatternField.Control.Align(), $oPatternField.Control.VerticalAlign(), $oPatternField.Control.BackgroundColor(), _
				$oPatternField.Control.Border(), $oPatternField.Control.BorderColor(), $oPatternField.Control.HideInactiveSelection(), $oPatternField.Control.Tag(), _
				$oPatternField.Control.HelpText(), $oPatternField.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPatternField.Control.Name = $sName
		$iError = ($oPatternField.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oLabelField = Default) Then
		$oPatternField.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPatternField.Control.LabelControl = $oLabelField.Control()
		$iError = ($oPatternField.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oPatternField.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPatternField.Control.WritingMode = $iTxtDir
		$iError = ($oPatternField.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iMaxLen = Default) Then
		$oPatternField.Control.setPropertyToDefault("MaxTextLen")

	ElseIf ($iMaxLen <> Null) Then
		If Not __LO_IntIsBetween($iMaxLen, -1, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPatternField.Control.MaxTextLen = $iMaxLen
		$iError = ($oPatternField.Control.MaxTextLen = $iMaxLen) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sEditMask = Default) Then
		$oPatternField.Control.setPropertyToDefault("EditMask")

	ElseIf ($sEditMask <> Null) Then
		If Not IsString($sEditMask) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oPatternField.Control.EditMask = $sEditMask
		$iError = ($oPatternField.Control.EditMask() = $sEditMask) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($sLiteralMask = Default) Then
		$oPatternField.Control.setPropertyToDefault("LiteralMask")

	ElseIf ($sLiteralMask <> Null) Then
		If Not IsString($sLiteralMask) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oPatternField.Control.LiteralMask = $sLiteralMask
		$iError = ($oPatternField.Control.LiteralMask() = $sLiteralMask) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bStrict = Default) Then
		$oPatternField.Control.setPropertyToDefault("StrictFormat")

	ElseIf ($bStrict <> Null) Then
		If Not IsBool($bStrict) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oPatternField.Control.StrictFormat = $bStrict
		$iError = ($oPatternField.Control.StrictFormat() = $bStrict) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bEnabled = Default) Then
		$oPatternField.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oPatternField.Control.Enabled = $bEnabled
		$iError = ($oPatternField.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bVisible = Default) Then
		$oPatternField.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oPatternField.Control.EnableVisible = $bVisible
		$iError = ($oPatternField.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bReadOnly = Default) Then
		$oPatternField.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oPatternField.Control.ReadOnly = $bReadOnly
		$iError = ($oPatternField.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($bPrintable = Default) Then
		$oPatternField.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oPatternField.Control.Printable = $bPrintable
		$iError = ($oPatternField.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($iMouseScroll = Default) Then
		$oPatternField.Control.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oPatternField.Control.MouseWheelBehavior = $iMouseScroll
		$iError = ($oPatternField.Control.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($bTabStop = Default) Then
		$oPatternField.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oPatternField.Control.Tabstop = $bTabStop
		$iError = ($oPatternField.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 8192) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oPatternField.Control.TabIndex = $iTabOrder
		$iError = ($oPatternField.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($sDefaultTxt = Default) Then
		$iError = BitOR($iError, 16384) ; Can't Default DefaultText.

	ElseIf ($sDefaultTxt <> Null) Then
		If Not IsString($sDefaultTxt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oPatternField.Control.DefaultText = $sDefaultTxt
		$iError = ($oPatternField.Control.DefaultText() = $sDefaultTxt) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 32768) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		__LOWriter_FormConSetGetFontDesc($oPatternField, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($iAlign = Default) Then
		$oPatternField.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oPatternField.Control.Align = $iAlign
		$iError = ($oPatternField.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($iVertAlign = Default) Then
		$oPatternField.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oPatternField.Control.VerticalAlign = $iVertAlign
		$iError = ($oPatternField.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($iBackColor = Default) Then
		$oPatternField.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oPatternField.Control.BackgroundColor = $iBackColor
		$iError = ($oPatternField.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($iBorder = Default) Then
		$oPatternField.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oPatternField.Control.Border = $iBorder
		$iError = ($oPatternField.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($iBorderColor = Default) Then
		$oPatternField.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oPatternField.Control.BorderColor = $iBorderColor
		$iError = ($oPatternField.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($bHideSel = Default) Then
		$oPatternField.Control.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oPatternField.Control.HideInactiveSelection = $bHideSel
		$iError = ($oPatternField.Control.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 4194304) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, 0)

		$oPatternField.Control.Tag = $sAddInfo
		$iError = ($oPatternField.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	If ($sHelpText = Default) Then
		$oPatternField.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 27, 0)

		$oPatternField.Control.HelpText = $sHelpText
		$iError = ($oPatternField.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 8388608))
	EndIf

	If ($sHelpURL = Default) Then
		$oPatternField.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 28, 0)

		$oPatternField.Control.HelpURL = $sHelpURL
		$iError = ($oPatternField.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 16777216))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConPatternFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConPatternFieldValue
; Description ...: Set or retrieve the current Pattern field value.
; Syntax ........: _LOWriter_FormConPatternFieldValue(ByRef $oPatternField[, $sValue = Null])
; Parameters ....: $oPatternField       - [in/out] an object. A Pattern Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sValue              - [optional] a string value. Default is Null. The value to set the field to.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPatternField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oPatternField not a Pattern Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sValue not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the current value of the control.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sValue
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning current value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current value.
;                  Call $sValue with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConPatternFieldGeneral, _LOWriter_FormConPatternFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConPatternFieldValue(ByRef $oPatternField, $sValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $sCurValue

	If Not IsObj($oPatternField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oPatternField) <> $LOW_FORM_CON_TYPE_PATTERN_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sValue) Then
		$sCurValue = $oPatternField.Control.Text()
		If Not IsString($sCurValue) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurValue)
	EndIf

	If ($sValue = Default) Then
		$oPatternField.Control.setPropertyToDefault("Text")

	Else
		If Not IsString($sValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPatternField.Control.Text = $sValue
		$iError = ($oPatternField.Control.Text() = $sValue) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConPatternFieldValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConPosition
; Description ...: Set or Retrieve the Control's position settings.
; Syntax ........: _LOWriter_FormConPosition(ByRef $oControl[, $iX = Null[, $iY = Null[, $iAnchor = Null[, $bProtectPos = Null]]]])
; Parameters ....: $oControl            - [in/out] an object. A Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iX                  - [optional] an integer value. Default is Null. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - [optional] an integer value. Default is Null. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iAnchor             - [optional] an integer value(0-4). Default is Null. The anchoring position for the Control. See Constants, $LOW_ANCHOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bProtectPos         - [optional] a boolean value. Default is Null. If True, the Shape's position is locked.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iAnchor not an Integer, less than 0 or greater than 4. See Constants, $LOW_ANCHOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bProtectPos not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Control's Position Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iX
;                  |                               2 = Error setting $iY
;                  |                               4 = Error setting $iAnchor
;                  |                               8 = Error setting $bProtectPos
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LO_UnitConvert, _LOWriter_FormConSize
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConPosition(ByRef $oControl, $iX = Null, $iY = Null, $iAnchor = Null, $bProtectPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPosition[4]
	Local $tPos

	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tPos = $oControl.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iX, $iY, $iAnchor, $bProtectPos) Then
		__LO_ArrayFill($avPosition, $tPos.X(), $tPos.Y(), $oControl.AnchorType(), $oControl.MoveProtect())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPosition)
	EndIf

	If ($iX <> Null) Or ($iY <> Null) Then
		If ($iX <> Null) Then
			If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

			$tPos.X = $iX
		EndIf

		If ($iY <> Null) Then
			If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$tPos.Y = $iY
		EndIf

		$oControl.Position = $tPos

		$iError = (__LO_VarsAreNull($iX)) ? ($iError) : ((__LO_IntIsBetween($oControl.Position.X(), $iX - 1, $iX + 1)) ? ($iError) : (BitOR($iError, 1)))
		$iError = (__LO_VarsAreNull($iY)) ? ($iError) : ((__LO_IntIsBetween($oControl.Position.Y(), $iY - 1, $iY + 1)) ? ($iError) : (BitOR($iError, 2)))
	EndIf

	If ($iAnchor <> Null) Then
		If Not __LO_IntIsBetween($iAnchor, $LOW_ANCHOR_AT_PARAGRAPH, $LOW_ANCHOR_AT_CHARACTER) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oControl.AnchorType = $iAnchor
		$iError = ($oControl.AnchorType() = $iAnchor) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bProtectPos <> Null) Then
		If Not IsBool($bProtectPos) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oControl.MoveProtect = $bProtectPos
		$iError = ($oControl.MoveProtect() = $bProtectPos) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConPushButtonGeneral
; Description ...: Set or Retrieve general Push Button properties.
; Syntax ........: _LOWriter_FormConPushButtonGeneral(ByRef $oPushButton[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bVisible = Null[, $bPrintable = Null[, $bTabStop = Null[, $iTabOrder = Null[, $bRepeat = Null[, $iDelay = Null[, $bTakeFocus = Null[, $bToggle = Null[, $iDefaultState = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $bWordBreak = Null[, $iAction = Null[, $sURL = Null[, $sFrame = Null[, $bDefault = Null[, $sGraphics = Null[, $iGraphicAlign = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oPushButton         - [in/out] an object. A Push Button Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $bTakeFocus          - [optional] a boolean value. Default is Null. If True, the button takes focus on clicking it.
;                  $bToggle             - [optional] a boolean value. Default is Null. If True, the button behaves like a toggle.
;                  $iDefaultState       - [optional] an integer value (0-1). Default is Null. The Default state of the Option button. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bWordBreak          - [optional] a boolean value. Default is Null. If True, line breaks are allowed to be used.
;                  $iAction             - [optional] an integer value (0-12). Default is Null. The action that occurs when the button is pushed. See Constants $LOW_FORM_CON_PUSH_CMD_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sURL                - [optional] a string value. Default is Null. The URL or Document path to open.
;                  $sFrame              - [optional] a string value. Default is Null. The frame to open the URL in. See Constants, $LOW_FRAME_TARGET_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bDefault            - [optional] a boolean value. Default is Null. If True, this will be set as the Default button.
;                  $sGraphics           - [optional] a string value. Default is Null. The path to an Image file.
;                  $iGraphicAlign       - [optional] an integer value (0-12). Default is Null. The Alignment of the Image. See Constants $LOW_FORM_CON_IMG_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPushButton not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oPushButton not a Push Button Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 11 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 13 Return 0 = $bTakeFocus not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $bToggle not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $iDefaultState not an Integer, less than 0 or greater than 1. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 16 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 17 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 18 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 19 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 20 Return 0 = $bWordBreak not a Boolean.
;                  @Error 1 @Extended 21 Return 0 = $iAction not an Integer, less than 0 or greater than 12. See Constants $LOW_FORM_CON_PUSH_CMD_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 22 Return 0 = $sURL not a String.
;                  @Error 1 @Extended 23 Return 0 = $sFrame not a String.
;                  @Error 1 @Extended 24 Return 0 = $sFrame not called with correct constant. See Constants, $LOW_FRAME_TARGET_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 25 Return 0 = $bDefault not a Boolean.
;                  @Error 1 @Extended 26 Return 0 = $sGraphics not a String.
;                  @Error 1 @Extended 27 Return 0 = $iGraphicAlign not an Integer, less than 0 or greater than 12. See Constants $LOW_FORM_CON_IMG_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 28 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 29 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 30 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bEnabled
;                  |                               16 = Error setting $bVisible
;                  |                               32 = Error setting $bPrintable
;                  |                               64 = Error setting $bTabStop
;                  |                               128 = Error setting $iTabOrder
;                  |                               256 = Error setting $bRepeat
;                  |                               512 = Error setting $iDelay
;                  |                               1024 = Error setting $bTakeFocus
;                  |                               2048 = Error setting $bToggle
;                  |                               4096 = Error setting $iDefaultState
;                  |                               8192 = Error setting $mFont
;                  |                               16384 = Error setting $iAlign
;                  |                               32768 = Error setting $iVertAlign
;                  |                               65536 = Error setting $iBackColor
;                  |                               131072 = Error setting $bWordBreak
;                  |                               262144 = Error setting $iAction
;                  |                               524288 = Error setting $sURL
;                  |                               1048576 = Error setting $sFrame
;                  |                               2097152 = Error setting $bDefault
;                  |                               4194304 = Error setting $sGraphics
;                  |                               8388608 = Error setting $iGraphicAlign
;                  |                               16777216 = Error setting $sAddInfo
;                  |                               33554432 = Error setting $sHelpText
;                  |                               67108864 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 27 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $sGraphics is called with an invalid Graphic URL, graphic is set to Null. The Return for $sGraphics is an Image Object.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $iDefaultState, $mFont, $sAddInfo.
; Related .......: _LOWriter_FormConPushButtonState
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConPushButtonGeneral(ByRef $oPushButton, $sName = Null, $sLabel = Null, $iTxtDir = Null, $bEnabled = Null, $bVisible = Null, $bPrintable = Null, $bTabStop = Null, $iTabOrder = Null, $bRepeat = Null, $iDelay = Null, $bTakeFocus = Null, $bToggle = Null, $iDefaultState = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $bWordBreak = Null, $iAction = Null, $sURL = Null, $sFrame = Null, $bDefault = Null, $sGraphics = Null, $iGraphicAlign = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iBtnAction
	Local Const $__LOW_PUSH_BTN_CMND_PUSH = 0, $__LOW_PUSH_BTN_CMND_SUBMIT = 1, $__LOW_PUSH_BTN_CMND_RESET = 2, $__LOW_PUSH_BTN_CMND_URL = 3 ; com.sun.star.form.FormButtonType
	Local $avControl[27]
	Local $asActions[13]

	If Not IsObj($oPushButton) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oPushButton) <> $LOW_FORM_CON_TYPE_PUSH_BUTTON) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asActions[$LOW_FORM_CON_PUSH_CMD_NONE] = ""
	$asActions[$LOW_FORM_CON_PUSH_CMD_SUBMIT_FORM] = ""
	$asActions[$LOW_FORM_CON_PUSH_CMD_RESET_FORM] = ""
	$asActions[$LOW_FORM_CON_PUSH_CMD_OPEN] = ""
	$asActions[$LOW_FORM_CON_PUSH_CMD_FIRST_REC] = ".uno:FormController/moveToFirst"
	$asActions[$LOW_FORM_CON_PUSH_CMD_LAST_REC] = ".uno:FormController/moveToLast"
	$asActions[$LOW_FORM_CON_PUSH_CMD_NEXT_REC] = ".uno:FormController/moveToNext"
	$asActions[$LOW_FORM_CON_PUSH_CMD_PREV_REC] = ".uno:FormController/moveToPrev"
	$asActions[$LOW_FORM_CON_PUSH_CMD_SAVE_REC] = ".uno:FormController/saveRecord"
	$asActions[$LOW_FORM_CON_PUSH_CMD_UNDO] = ".uno:FormController/undoRecord"
	$asActions[$LOW_FORM_CON_PUSH_CMD_NEW_REC] = ".uno:FormController/moveToNew"
	$asActions[$LOW_FORM_CON_PUSH_CMD_DELETE_REC] = ".uno:FormController/deleteRecord"
	$asActions[$LOW_FORM_CON_PUSH_CMD_REFRESH_FORM] = ".uno:FormController/refreshForm"

	Switch $oPushButton.Control.ButtonType()
		Case $__LOW_PUSH_BTN_CMND_PUSH
			$iBtnAction = $LOW_FORM_CON_PUSH_CMD_NONE

		Case $__LOW_PUSH_BTN_CMND_SUBMIT
			$iBtnAction = $LOW_FORM_CON_PUSH_CMD_SUBMIT_FORM

		Case $__LOW_PUSH_BTN_CMND_RESET
			$iBtnAction = $LOW_FORM_CON_PUSH_CMD_RESET_FORM

		Case $__LOW_PUSH_BTN_CMND_URL
			$iBtnAction = $LOW_FORM_CON_PUSH_CMD_OPEN

			If StringInStr($oPushButton.Control.TargetURL(), ".uno:FormController/") Then
				Switch $oPushButton.Control.TargetURL()
					Case ".uno:FormController/moveToFirst"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_FIRST_REC

					Case ".uno:FormController/moveToLast"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_LAST_REC

					Case ".uno:FormController/moveToNext"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_NEXT_REC

					Case ".uno:FormController/moveToPrev"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_PREV_REC

					Case ".uno:FormController/saveRecord"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_SAVE_REC

					Case ".uno:FormController/undoRecord"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_UNDO

					Case ".uno:FormController/moveToNew"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_NEW_REC

					Case ".uno:FormController/deleteRecord"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_DELETE_REC

					Case ".uno:FormController/refreshForm"
						$iBtnAction = $LOW_FORM_CON_PUSH_CMD_REFRESH_FORM
				EndSwitch
			EndIf
	EndSwitch

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $bEnabled, $bVisible, $bPrintable, $bTabStop, $iTabOrder, $bRepeat, $iDelay, $bTakeFocus, $bToggle, $iDefaultState, $mFont, $iAlign, $iVertAlign, $iBackColor, $bWordBreak, $iAction, $sURL, $sFrame, $bDefault, $sGraphics, $iGraphicAlign, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oPushButton.Control.Name(), $oPushButton.Control.Label(), $oPushButton.Control.WritingMode(), $oPushButton.Control.Enabled(), _
				$oPushButton.Control.EnableVisible(), $oPushButton.Control.Printable(), $oPushButton.Control.Tabstop(), $oPushButton.Control.TabIndex(), _
				$oPushButton.Control.Repeat(), $oPushButton.Control.RepeatDelay(), $oPushButton.Control.FocusOnClick(), $oPushButton.Control.Toggle(), _
				$oPushButton.Control.DefaultState(), __LOWriter_FormConSetGetFontDesc($oPushButton), $oPushButton.Control.Align(), $oPushButton.Control.VerticalAlign(), _
				$oPushButton.Control.BackgroundColor(), $oPushButton.Control.MultiLine(), $iBtnAction, $oPushButton.Control.TargetURL(), $oPushButton.Control.TargetFrame(), _
				$oPushButton.Control.DefaultButton(), $oPushButton.Control.Graphic(), $oPushButton.Control.ImagePosition(), $oPushButton.Control.Tag(), _
				$oPushButton.Control.HelpText(), $oPushButton.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPushButton.Control.Name = $sName
		$iError = ($oPushButton.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oPushButton.Control.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPushButton.Control.Label = $sLabel
		$iError = ($oPushButton.Control.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oPushButton.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPushButton.Control.WritingMode = $iTxtDir
		$iError = ($oPushButton.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bEnabled = Default) Then
		$oPushButton.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPushButton.Control.Enabled = $bEnabled
		$iError = ($oPushButton.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bVisible = Default) Then
		$oPushButton.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPushButton.Control.EnableVisible = $bVisible
		$iError = ($oPushButton.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bPrintable = Default) Then
		$oPushButton.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oPushButton.Control.Printable = $bPrintable
		$iError = ($oPushButton.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bTabStop = Default) Then
		$oPushButton.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oPushButton.Control.Tabstop = $bTabStop
		$iError = ($oPushButton.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 128) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oPushButton.Control.TabIndex = $iTabOrder
		$iError = ($oPushButton.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bRepeat = Default) Then
		$oPushButton.Control.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oPushButton.Control.Repeat = $bRepeat
		$iError = ($oPushButton.Control.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iDelay = Default) Then
		$oPushButton.Control.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oPushButton.Control.RepeatDelay = $iDelay
		$iError = ($oPushButton.Control.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($bTakeFocus = Default) Then
		$oPushButton.Control.setPropertyToDefault("FocusOnClick")

	ElseIf ($bTakeFocus <> Null) Then
		If Not IsBool($bTakeFocus) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oPushButton.Control.FocusOnClick = $bTakeFocus
		$iError = ($oPushButton.Control.FocusOnClick() = $bTakeFocus) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($bToggle = Default) Then
		$oPushButton.Control.setPropertyToDefault("Toggle")

	ElseIf ($bToggle <> Null) Then
		If Not IsBool($bToggle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oPushButton.Control.Toggle = $bToggle
		$iError = ($oPushButton.Control.Toggle() = $bToggle) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($iDefaultState = Default) Then
		$iError = BitOR($iError, 4096) ; Can't Default DefaultState.

	ElseIf ($iDefaultState <> Null) Then
		If Not __LO_IntIsBetween($iDefaultState, $LOW_FORM_CON_CHKBX_STATE_NOT_SELECTED, $LOW_FORM_CON_CHKBX_STATE_SELECTED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oPushButton.Control.DefaultState = $iDefaultState
		$iError = ($oPushButton.Control.DefaultState() = $iDefaultState) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 8192) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		__LOWriter_FormConSetGetFontDesc($oPushButton, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iAlign = Default) Then
		$oPushButton.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oPushButton.Control.Align = $iAlign
		$iError = ($oPushButton.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iVertAlign = Default) Then
		$oPushButton.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oPushButton.Control.VerticalAlign = $iVertAlign
		$iError = ($oPushButton.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($iBackColor = Default) Then
		$oPushButton.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oPushButton.Control.BackgroundColor = $iBackColor
		$iError = ($oPushButton.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($bWordBreak = Default) Then
		$oPushButton.Control.setPropertyToDefault("MultiLine")

	ElseIf ($bWordBreak <> Null) Then
		If Not IsBool($bWordBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oPushButton.Control.MultiLine = $bWordBreak
		$iError = ($oPushButton.Control.MultiLine() = $bWordBreak) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($iAction = Default) Then
		$oPushButton.Control.setPropertyToDefault("ButtonType")

		Switch $iBtnAction
			Case $LOW_FORM_CON_PUSH_CMD_NONE, $LOW_FORM_CON_PUSH_CMD_SUBMIT_FORM, $LOW_FORM_CON_PUSH_CMD_RESET_FORM
				$oPushButton.Control.setPropertyToDefault("TargetURL")

			Case $LOW_FORM_CON_PUSH_CMD_OPEN

			Case $LOW_FORM_CON_PUSH_CMD_FIRST_REC, $LOW_FORM_CON_PUSH_CMD_LAST_REC, $LOW_FORM_CON_PUSH_CMD_NEXT_REC, $LOW_FORM_CON_PUSH_CMD_PREV_REC, _
					$LOW_FORM_CON_PUSH_CMD_SAVE_REC, $LOW_FORM_CON_PUSH_CMD_UNDO, $LOW_FORM_CON_PUSH_CMD_NEW_REC, $LOW_FORM_CON_PUSH_CMD_DELETE_REC, _
					$LOW_FORM_CON_PUSH_CMD_REFRESH_FORM
				$oPushButton.Control.setPropertyToDefault("TargetURL")
		EndSwitch

	ElseIf ($iAction <> Null) Then
		If Not __LO_IntIsBetween($iAction, $LOW_FORM_CON_PUSH_CMD_NONE, $LOW_FORM_CON_PUSH_CMD_REFRESH_FORM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		Switch $iAction
			Case $LOW_FORM_CON_PUSH_CMD_NONE
				$oPushButton.Control.ButtonType = $__LOW_PUSH_BTN_CMND_PUSH
				$sURL = $asActions[$iAction]
				$iError = (($oPushButton.Control.ButtonType() = $__LOW_PUSH_BTN_CMND_PUSH)) ? ($iError) : (BitOR($iError, 262144))

			Case $LOW_FORM_CON_PUSH_CMD_SUBMIT_FORM
				$oPushButton.Control.ButtonType = $__LOW_PUSH_BTN_CMND_SUBMIT
				$sURL = $asActions[$iAction]
				$iError = (($oPushButton.Control.ButtonType() = $__LOW_PUSH_BTN_CMND_SUBMIT)) ? ($iError) : (BitOR($iError, 262144))

			Case $LOW_FORM_CON_PUSH_CMD_RESET_FORM
				$oPushButton.Control.ButtonType = $__LOW_PUSH_BTN_CMND_RESET
				$sURL = $asActions[$iAction]
				$iError = (($oPushButton.Control.ButtonType() = $__LOW_PUSH_BTN_CMND_RESET)) ? ($iError) : (BitOR($iError, 262144))

			Case $LOW_FORM_CON_PUSH_CMD_OPEN
				$oPushButton.Control.ButtonType = $__LOW_PUSH_BTN_CMND_URL
				If ($iBtnAction <> $LOW_FORM_CON_PUSH_CMD_OPEN) And ($sURL = Null) Then $sURL = $asActions[$iAction]
				$iError = (($oPushButton.Control.ButtonType() = $__LOW_PUSH_BTN_CMND_URL)) ? ($iError) : (BitOR($iError, 262144))

			Case $LOW_FORM_CON_PUSH_CMD_FIRST_REC, $LOW_FORM_CON_PUSH_CMD_LAST_REC, $LOW_FORM_CON_PUSH_CMD_NEXT_REC, $LOW_FORM_CON_PUSH_CMD_PREV_REC, _
					$LOW_FORM_CON_PUSH_CMD_SAVE_REC, $LOW_FORM_CON_PUSH_CMD_UNDO, $LOW_FORM_CON_PUSH_CMD_NEW_REC, $LOW_FORM_CON_PUSH_CMD_DELETE_REC, _
					$LOW_FORM_CON_PUSH_CMD_REFRESH_FORM
				$oPushButton.Control.ButtonType = $__LOW_PUSH_BTN_CMND_URL
				$sURL = $asActions[$iAction]
				$iError = (($oPushButton.Control.ButtonType() = $__LOW_PUSH_BTN_CMND_URL)) ? ($iError) : (BitOR($iError, 262144))
		EndSwitch
	EndIf

	If ($sURL = Default) Then
		$oPushButton.Control.setPropertyToDefault("TargetURL")

	ElseIf ($sURL <> Null) Then
		If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oPushButton.Control.TargetURL = $sURL
		$iError = ($oPushButton.Control.TargetURL() = $sURL) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($sFrame = Default) Then
		$oPushButton.Control.setPropertyToDefault("TargetFrame")

	ElseIf ($sFrame <> Null) Then
		If Not IsString($sFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)
		If ($sFrame <> $LOW_FRAME_TARGET_TOP) And _
				($sFrame <> $LOW_FRAME_TARGET_PARENT) And _
				($sFrame <> $LOW_FRAME_TARGET_BLANK) And _
				($sFrame <> $LOW_FRAME_TARGET_SELF) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)
		$oPushButton.Control.TargetFrame = $sFrame
		$iError = ($oPushButton.Control.TargetFrame() = $sFrame) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($bDefault = Default) Then
		$oPushButton.Control.setPropertyToDefault("DefaultButton")

	ElseIf ($bDefault <> Null) Then
		If Not IsBool($bDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oPushButton.Control.DefaultButton = $bDefault
		$iError = ($oPushButton.Control.DefaultButton() = $bDefault) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($sGraphics = Default) Then
		$oPushButton.Control.setPropertyToDefault("ImageURL")
		$oPushButton.Control.setPropertyToDefault("Graphic")

	ElseIf ($sGraphics <> Null) Then
		If Not IsString($sGraphics) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, 0)

		$oPushButton.Control.ImageURL = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)
		$iError = ($oPushButton.Control.ImageURL() = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	If ($iGraphicAlign = Default) Then
		$oPushButton.Control.setPropertyToDefault("ImagePosition")

	ElseIf ($iGraphicAlign <> Null) Then
		If Not __LO_IntIsBetween($iGraphicAlign, $LOW_FORM_CON_IMG_ALIGN_LEFT_TOP, $LOW_FORM_CON_IMG_ALIGN_CENTERED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 27, 0)

		$oPushButton.Control.ImagePosition = $iGraphicAlign
		$iError = ($oPushButton.Control.ImagePosition() = $iGraphicAlign) ? ($iError) : (BitOR($iError, 8388608))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 16777216) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 28, 0)

		$oPushButton.Control.Tag = $sAddInfo
		$iError = ($oPushButton.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 16777216))
	EndIf

	If ($sHelpText = Default) Then
		$oPushButton.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 29, 0)

		$oPushButton.Control.HelpText = $sHelpText
		$iError = ($oPushButton.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 33554432))
	EndIf

	If ($sHelpURL = Default) Then
		$oPushButton.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 30, 0)

		$oPushButton.Control.HelpURL = $sHelpURL
		$iError = ($oPushButton.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 67108864))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConPushButtonGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConPushButtonState
; Description ...: Set or Retrieve the current Push Button state.
; Syntax ........: _LOWriter_FormConPushButtonState(ByRef $oPushButton[, $iState = Null])
; Parameters ....: $oPushButton         - [in/out] an object. A Push Button Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iState              - [optional] an integer value (0-1). Default is Null. The state of the Push Button. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPushButton not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oPushButton not a Push Button Control.
;                  @Error 1 @Extended 3 Return 0 = $iState not an Integer, less than 0 or greater than 1. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current control State.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iState
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current Push Button State as an Integer, matching one of the constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current Push Button state.
;                  Setting the state to selected DOES NOT simulate clicking the button.
;                  The Push button State is only valid when Toggle is active.
;                  Call $iState with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConPushButtonGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConPushButtonState(ByRef $oPushButton, $iState = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurState

	If Not IsObj($oPushButton) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oPushButton) <> $LOW_FORM_CON_TYPE_PUSH_BUTTON) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iState) Then
		$iCurState = $oPushButton.Control.State()
		If Not IsInt($iCurState) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurState)
	EndIf

	If ($iState = Default) Then
		$oPushButton.Control.setPropertyToDefault("State")

	Else
		If Not __LO_IntIsBetween($iState, $LOW_FORM_CON_CHKBX_STATE_NOT_SELECTED, $LOW_FORM_CON_CHKBX_STATE_NOT_DEFINED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPushButton.Control.State = $iState
		$iError = ($oPushButton.Control.State() = $iState) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConPushButtonState

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConsGetList
; Description ...: Retrieve an array of Control Objects contained in a Document or a Form.
; Syntax ........: _LOWriter_FormConsGetList(ByRef $oObj[, $iType = $LOW_FORM_CON_TYPE_ALL])
; Parameters ....: $oObj                - [in/out] an object. Either a Document Object or a Form object, or a Grouped Control. See Remarks. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function, or a Form Object returned from a previous _LOWriter_FormsGetList, or _LOWriter_FormAdd function. Also a Grouped Control or Group Box returned from a _LOWriter_FormConsGetList function.
;                  $iType               - [optional] an integer value (1-1048575). Default is $LOW_FORM_CON_TYPE_ALL. The type of control(s) to return in the array. Can be BitOr'd together. See Constants $LOW_FORM_CON_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 1 or greater than 1048575. See Constants $LOW_FORM_CON_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = Called Object in $oObj, not a Document Object, not a Form Object, and not a Grouped Control.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify parent document of Form.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Draw Page object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Shape Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Control Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning a 2D array of Control Objects in the first column, and the type of Control in the second column, corresponding to the Constants $LOW_FORM_CON_* as defined in LibreOfficeWriter_Constants.au3
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If a Document object is called in $oObj, all the controls are returned (except controls in a Grouped Control). If a Form Object is called in $oObj, only the controls contained in the Form are returned. And if a Grouped control is called, only controls in the group are returned.
;                  If there is a Grouped Control (a group containing a Group Box, and usually an option button) present, its object will be returned with the appropriate Constant, you can call this function with its object to obtain the controls grouped in the group box.
;                  Currently I am only able test a single layer Grouped control, as trying to nest Grouped controls crashes my LibreOffice.
; Related .......: _LOWriter_FormConInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConsGetList(ByRef $oObj, $iType = $LOW_FORM_CON_TYPE_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aoControls[0][2]
	Local $oShapes, $oShape, $oDoc, $oControl
	Local $iCount = 0, $iControlType

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iType, $LOW_FORM_CON_TYPE_CHECK_BOX, $LOW_FORM_CON_TYPE_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oObj.supportsService("com.sun.star.form.component.Form") Then ; Get only controls in the form.

		$oDoc = $oObj ; Identify the parent document.

		Do
			$oDoc = $oDoc.getParent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		Until $oDoc.supportsService("com.sun.star.text.TextDocument")

		$oShapes = $oDoc.DrawPage()
		If Not IsObj($oShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		ReDim $aoControls[$oShapes.Count()][2]

		For $i = 0 To $oShapes.Count() - 1
			$oShape = $oShapes.getByIndex($i)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			If $oShape.supportsService("com.sun.star.drawing.ControlShape") And ($oShape.Control.Parent() = $oObj) Then ; If shape is a single control, and is contained in the form.

				$iControlType = __LOWriter_FormConIdentify($oShape)
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

				If BitAND($iType, $iControlType) Then
					$aoControls[$iCount][0] = $oShape
					$aoControls[$iCount][1] = $iControlType
					$iCount += 1
				EndIf

			ElseIf $oShape.supportsService("com.sun.star.drawing.GroupShape") And ($oShape.getByIndex(0).Control.Parent() = $oObj) Then ; If shape is a group control, and the first control contained in it is contained in the form.
				$iControlType = __LOWriter_FormConIdentify($oShape)
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

				If BitAND($iType, $iControlType) Then
					$aoControls[$iCount][0] = $oShape
					$aoControls[$iCount][1] = $LOW_FORM_CON_TYPE_GROUPED_CONTROL

					$iCount += 1
				EndIf
			EndIf
			Sleep((IsInt(($i / $__LOWCONST_SLEEP_DIV))) ? (10) : (0))
		Next

	ElseIf $oObj.supportsService("com.sun.star.text.TextDocument") Then ; Get all controls in document.
		$oShapes = $oObj.DrawPage()
		If Not IsObj($oShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		ReDim $aoControls[$oShapes.Count()][2]

		For $i = 0 To $oShapes.Count() - 1
			$oShape = $oShapes.getByIndex($i)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			If $oShape.supportsService("com.sun.star.drawing.ControlShape") Or $oShape.supportsService("com.sun.star.drawing.GroupShape") Then
				$iControlType = __LOWriter_FormConIdentify($oShape)
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

				If BitAND($iType, $iControlType) Then
					$aoControls[$iCount][0] = $oShape
					$aoControls[$iCount][1] = $iControlType
					$iCount += 1
				EndIf
			EndIf
			Sleep((IsInt(($i / $__LOWCONST_SLEEP_DIV))) ? (10) : (0))
		Next

	ElseIf $oObj.supportsService("com.sun.star.drawing.GroupShape") Then
		ReDim $aoControls[$oObj.Count()][2]

		For $i = 0 To $oObj.Count() - 1
			$oControl = $oObj.getByIndex($i)
			If Not IsObj($oControl) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

			If $oControl.supportsService("com.sun.star.drawing.ControlShape") Then
				$iControlType = __LOWriter_FormConIdentify($oControl)
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

				If BitAND($iType, $iControlType) Then
					$aoControls[$iCount][0] = $oControl
					$aoControls[$iCount][1] = $iControlType
					$iCount += 1
				EndIf
			EndIf
			Sleep((IsInt(($i / $__LOWCONST_SLEEP_DIV))) ? (10) : (0))
		Next

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; wrong type of input item.
	EndIf

	ReDim $aoControls[$iCount][2]

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoControls)
EndFunc   ;==>_LOWriter_FormConsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConSize
; Description ...: Set or Retrieve Control Size related settings.
; Syntax ........: _LOWriter_FormConSize(ByRef $oControl[, $iWidth = Null[, $iHeight = Null[, $bProtectSize = Null]]])
; Parameters ....: $oControl            - [in/out] an object. A Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the Shape, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the Shape, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $bProtectSize        - [optional] a boolean value. Default is Null. If True, Locks the size of the Shape.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer, or less than 51.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer, or less than 51.
;                  @Error 1 @Extended 4 Return 0 = $bProtectSize not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Control Size Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iHeight
;                  |                               4 = Error setting $bProtectSize
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  I have skipped Keep Ratio, as currently it seems unable to be set for controls.
; Related .......: _LO_UnitConvert, _LOWriter_FormConPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConSize(ByRef $oControl, $iWidth = Null, $iHeight = Null, $bProtectSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSize[3]
	Local $tSize

	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tSize = $oControl.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iHeight, $bProtectSize) Then
		__LO_ArrayFill($avSize, $tSize.Width(), $tSize.Height(), $oControl.SizeProtect())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSize)
	EndIf

	If ($iWidth <> Null) Or ($iHeight <> Null) Then
		If ($iWidth <> Null) Then ; Min 51
			If Not __LO_IntIsBetween($iWidth, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

			$tSize.Width = $iWidth
		EndIf

		If ($iHeight <> Null) Then
			If Not __LO_IntIsBetween($iHeight, 51) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$tSize.Height = $iHeight
		EndIf

		$oControl.Size = $tSize

		$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : ((__LO_IntIsBetween($oControl.Size.Width(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
		$iError = (__LO_VarsAreNull($iHeight)) ? ($iError) : ((__LO_IntIsBetween($oControl.Size.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 2)))
	EndIf

	If ($bProtectSize <> Null) Then
		If Not IsBool($bProtectSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oControl.SizeProtect = $bProtectSize
		$iError = ($oControl.SizeProtect() = $bProtectSize) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConCheckBoxData
; Description ...: Set or Retrieve Table Control Check Box Data Properties.
; Syntax ........: _LOWriter_FormConTableConCheckBoxData(ByRef $oCheckBox[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oCheckBox           - [in/out] an object. A Table Control Check Box Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCheckBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oCheckBox not a Check Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
;                  Reference Values are not included here as they are applicable to Calc only, as far as I can ascertain.
; Related .......: _LOWriter_FormConCheckBoxGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConCheckBoxData(ByRef $oCheckBox, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oCheckBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oCheckBox) <> $LOW_FORM_CON_TYPE_CHECK_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oCheckBox.DataField(), $oCheckBox.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCheckBox.DataField = $sDataField
		$iError = ($oCheckBox.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCheckBox.InputRequired = $bInputRequired
		$iError = ($oCheckBox.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConCheckBoxData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConCheckBoxGeneral
; Description ...: Set or Retrieve general Table Control Checkbox control properties.
; Syntax ........: _LOWriter_FormConTableConCheckBoxGeneral(ByRef $oCheckBox[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $bEnabled = Null[, $iDefaultState = Null[, $iWidth = Null[, $iStyle = Null[, $iAlign = Null[, $bWordBreak = Null[, $bTriState = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]])
; Parameters ....: $oCheckBox           - [in/out] an object. A Table Control Checkbox Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $iDefaultState       - [optional] an integer value (0-2). Default is Null. The Default state of the Checkbox, $LOW_FORM_CON_CHKBX_STATE_NOT_DEFINED is only available if $bTriState is True. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iWidth              - [optional] an integer value (100-200000). Default is Null. The width of the Column tab, in Hundredths of a Millimeter (HMM).
;                  $iStyle              - [optional] an integer value (1-2). Default is Null. The display style of the checkbox. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bWordBreak          - [optional] a boolean value. Default is Null. If True, line breaks are allowed to be used.
;                  $bTriState           - [optional] a boolean value. Default is Null. If True, the checkbox will have a third checked state.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCheckBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oCheckBox not a Check Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iDefaultState not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_CHKBX_STATE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iWidth not an Integer, less than 100 or greater than 200,000.
;                  @Error 1 @Extended 9 Return 0 = $iStyle not an Integer, less than 1 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = $bWordBreak not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $bTriState not a Boolean.
;                  @Error 1 @Extended 13 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 14 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 15 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bEnabled
;                  |                               16 = Error setting $iDefaultState
;                  |                               32 = Error setting $iWidth
;                  |                               64 = Error setting $iStyle
;                  |                               128 = Error setting $iAlign
;                  |                               256 = Error setting $bWordBreak
;                  |                               512 = Error setting $bTriState
;                  |                               1024 = Error setting $sAddInfo
;                  |                               2048 = Error setting $sHelpText
;                  |                               4096 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 13 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iDefaultState, $sAddInfo.
; Related .......: _LOWriter_FormConCheckBoxData, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConCheckBoxGeneral(ByRef $oCheckBox, $sName = Null, $sLabel = Null, $iTxtDir = Null, $bEnabled = Null, $iDefaultState = Null, $iWidth = Null, $iStyle = Null, $iAlign = Null, $bWordBreak = Null, $bTriState = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[13]

	If Not IsObj($oCheckBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ( __LOWriter_FormConIdentify($oCheckBox) <> $LOW_FORM_CON_TYPE_CHECK_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $bEnabled, $iDefaultState, $iWidth, $iStyle, $iAlign, $bWordBreak, $bTriState, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oCheckBox.Name(), $oCheckBox.Label(), $oCheckBox.WritingMode(), $oCheckBox.Enabled(), _
				$oCheckBox.DefaultState(), Int($oCheckBox.Width() * 10), _ ; Multiply width by 10 to get the Hundredths of a Millimeter value.
				$oCheckBox.VisualEffect(), $oCheckBox.Align(), $oCheckBox.MultiLine(), $oCheckBox.TriState(), _
				$oCheckBox.Tag(), $oCheckBox.HelpText(), $oCheckBox.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCheckBox.Name = $sName
		$iError = ($oCheckBox.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oCheckBox.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCheckBox.Label = $sLabel
		$iError = ($oCheckBox.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oCheckBox.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oCheckBox.WritingMode = $iTxtDir
		$iError = ($oCheckBox.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bEnabled = Default) Then
		$oCheckBox.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oCheckBox.Enabled = $bEnabled
		$iError = ($oCheckBox.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iDefaultState = Default) Then
		$iError = BitOR($iError, 16) ; Can't Default DefaultState.

	ElseIf ($iDefaultState <> Null) Then
		If Not __LO_IntIsBetween($iDefaultState, $LOW_FORM_CON_CHKBX_STATE_NOT_SELECTED, $LOW_FORM_CON_CHKBX_STATE_NOT_DEFINED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oCheckBox.DefaultState = $iDefaultState
		$iError = ($oCheckBox.DefaultState() = $iDefaultState) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iWidth = Default) Then
		$oCheckBox.setPropertyToDefault("Width")

	ElseIf ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 100, 200000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oCheckBox.Width = Round($iWidth / 10) ; Divide the Hundredths of a Millimeter value by 10 to obtain 10th MM.
		$iError = ($oCheckBox.Width() = Round($iWidth / 10)) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iStyle = Default) Then
		$oCheckBox.setPropertyToDefault("VisualEffect")

	ElseIf ($iStyle <> Null) Then
		If Not __LO_IntIsBetween($iStyle, $LOW_FORM_CON_BORDER_3D, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oCheckBox.VisualEffect = $iStyle
		$iError = ($oCheckBox.VisualEffect() = $iStyle) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iAlign = Default) Then
		$oCheckBox.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oCheckBox.Align = $iAlign
		$iError = ($oCheckBox.Align() = $iAlign) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bWordBreak = Default) Then
		$oCheckBox.setPropertyToDefault("MultiLine")

	ElseIf ($bWordBreak <> Null) Then
		If Not IsBool($bWordBreak) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oCheckBox.MultiLine = $bWordBreak
		$iError = ($oCheckBox.MultiLine() = $bWordBreak) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bTriState = Default) Then
		$oCheckBox.setPropertyToDefault("TriState")

	ElseIf ($bTriState <> Null) Then
		If Not IsBool($bTriState) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oCheckBox.TriState = $bTriState
		$iError = ($oCheckBox.TriState = $bTriState) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oCheckBox.Tag = $sAddInfo
		$iError = ($oCheckBox.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($sHelpText = Default) Then
		$oCheckBox.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oCheckBox.HelpText = $sHelpText
		$iError = ($oCheckBox.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($sHelpURL = Default) Then
		$oCheckBox.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oCheckBox.HelpURL = $sHelpURL
		$iError = ($oCheckBox.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConCheckBoxGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConColumnAdd
; Description ...: Add a column to a Table Control.
; Syntax ........: _LOWriter_FormConTableConColumnAdd(ByRef $oTableCon, $iControl[, $iPos = Null])
; Parameters ....: $oTableCon           - [in/out] an object. A Table Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iControl            - an integer value. The control type to insert. See Constants $LOW_FORM_CON_TYPE_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
;                  $iPos                - [optional] an integer value. Default is Null. The position in the Column list to insert the new Column. 0 = insert at the beginning. Null = insert at the end.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTableCon not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTableCon not a Table Control Object.
;                  @Error 1 @Extended 3 Return 0 = Control type called in $iControl not an Integer, or not one of the accepted controls.
;                  @Error 1 @Extended 4 Return 0 = $iPos not an Integer, less than 0 or greater than count of Columns + 1.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Column object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Control type name.
;                  @Error 3 @Extended 3 Return 0 = Failed to split Control type name.
;                  @Error 3 @Extended 4 Return 0 = Failed to insert new Column.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully created the Column, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only the following Control types are allowed to be created as a column tab: $LOW_FORM_CON_TYPE_FORMATTED_FIELD, $LOW_FORM_CON_TYPE_LIST_BOX, $LOW_FORM_CON_TYPE_NUMERIC_FIELD, $LOW_FORM_CON_TYPE_PATTERN_FIELD, $LOW_FORM_CON_TYPE_TEXT_BOX, $LOW_FORM_CON_TYPE_TIME_FIELD
; Related .......: _LOWriter_FormConTableConColumnsGetList, _LOWriter_FormConTableConColumnDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConColumnAdd(ByRef $oTableCon, $iControl, $iPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumn
	Local $iCount = 1
	Local $sControl, $sName, $sAccepted = $LOW_FORM_CON_TYPE_FORMATTED_FIELD & ":" & $LOW_FORM_CON_TYPE_LIST_BOX & ":" & $LOW_FORM_CON_TYPE_NUMERIC_FIELD & ":" & $LOW_FORM_CON_TYPE_PATTERN_FIELD & ":" & $LOW_FORM_CON_TYPE_TEXT_BOX & ":" & $LOW_FORM_CON_TYPE_TIME_FIELD

	If Not IsObj($oTableCon) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTableCon) <> $LOW_FORM_CON_TYPE_TABLE_CONTROL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iControl, $LOW_FORM_CON_TYPE_CHECK_BOX, $LOW_FORM_CON_TYPE_DATE_FIELD, "", $sAccepted) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iPos = ($iPos = Null) ? ($oTableCon.Control.Count()) : ($iPos)

	If ($iPos <> Null) And Not __LO_IntIsBetween($iPos, 0, $oTableCon.Control.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$sControl = __LOWriter_FormConIdentify(Null, $iControl)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$sControl = StringTrimLeft($sControl, StringInStr($sControl, ".", 0, -1))
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oColumn = $oTableCon.Control.createColumn($sControl)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Switch $iControl
		Case $LOW_FORM_CON_TYPE_CHECK_BOX
			$sName = "Check Box "

		Case $LOW_FORM_CON_TYPE_COMBO_BOX
			$sName = "Combo Box "

		Case $LOW_FORM_CON_TYPE_CURRENCY_FIELD
			$sName = "Currency Field "

		Case $LOW_FORM_CON_TYPE_DATE_FIELD
			$sName = "Date Field "

		Case $LOW_FORM_CON_TYPE_FORMATTED_FIELD
			$sName = "Formatted Field "

		Case $LOW_FORM_CON_TYPE_LIST_BOX
			$sName = "List Box "

		Case $LOW_FORM_CON_TYPE_NUMERIC_FIELD
			$sName = "Numeric Field "

		Case $LOW_FORM_CON_TYPE_PATTERN_FIELD
			$sName = "Pattern Field "

		Case $LOW_FORM_CON_TYPE_TEXT_BOX
			$sName = "Text Box "

		Case $LOW_FORM_CON_TYPE_TIME_FIELD
			$sName = "Time Field "

		Case Else
			$sName = "Unknown Control "
	EndSwitch

	While $oTableCon.Control.hasByName($sName & $iCount)
		$iCount += 1
		Sleep((IsInt(($iCount / $__LOWCONST_SLEEP_DIV))) ? (10) : (0))
	WEnd

	$oColumn.Name = $sName & $iCount
	$oColumn.Label = $sName & $iCount

	$oTableCon.Control.insertByIndex($iPos, $oColumn)

	$oColumn = $oTableCon.Control.getByIndex($iPos)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOWriter_FormConTableConColumnAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConColumnDelete
; Description ...: Delete a Table Control Column.
; Syntax ........: _LOWriter_FormConTableConColumnDelete(ByRef $oColumn)
; Parameters ....: $oColumn             - [in/out] an object. A Column object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oColumn not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Column's parent.
;                  @Error 3 @Extended 2 Return 0 = Failed to delete the Column.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Column was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormConTableConColumnAdd, _LOWriter_FormConTableConColumnsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConColumnDelete(ByRef $oColumn)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oParent

	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oParent = $oColumn.Parent()
	If Not IsObj($oParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To $oParent.Count() - 1
		If ($oParent.getByIndex($i) = $oColumn) Then
			$oParent.removeByIndex($i) ; The name can be the same as another control, so I have to remove by Index.

			Return SetError($__LO_STATUS_SUCCESS, 0, 1)
		EndIf

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
EndFunc   ;==>_LOWriter_FormConTableConColumnDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConColumnsGetList
; Description ...: Retrieve a list of Columns contained in a Table Control.
; Syntax ........: _LOWriter_FormConTableConColumnsGetList(ByRef $oTableCon)
; Parameters ....: $oTableCon           - [in/out] an object. A Table Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTableCon not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTableCon not a Table Control Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve count of Columns.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Column object.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify Column type.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning a two column array containing Column objects. See remarks. @Extended is set to the number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The returned array will contain two columns. The first column ($aArray[0][0], contains the Column object, and the second column ($aArray[0][1], contains the Column type, corresponding to one of the constants $LOW_FORM_CON_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Related .......: _LOWriter_FormConTableConColumnAdd, _LOWriter_FormConTableConColumnDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConColumnsGetList(ByRef $oTableCon)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumn
	Local $iCount
	Local $avColumns[0][2]

	If Not IsObj($oTableCon) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTableCon) <> $LOW_FORM_CON_TYPE_TABLE_CONTROL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oTableCon.Control.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	ReDim $avColumns[$iCount][2]

	For $i = 0 To $iCount - 1
		$oColumn = $oTableCon.Control.getByIndex($i)
		If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$avColumns[$i][0] = $oColumn
		$avColumns[$i][1] = __LOWriter_FormConIdentify($oColumn)
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $avColumns)
EndFunc   ;==>_LOWriter_FormConTableConColumnsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConComboBoxData
; Description ...: Set or Retrieve Table Control Combo Box Data Properties.
; Syntax ........: _LOWriter_FormConTableConComboBoxData(ByRef $oComboBox[, $sDataField = Null[, $bEmptyIsNull = Null[, $bInputRequired = Null[, $iType = Null[, $sListContent = Null]]]]])
; Parameters ....: $oComboBox           - [in/out] an object. A Table Control Combo Box Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bEmptyIsNull        - [optional] a boolean value. Default is Null. If True, an empty string will be treated as a Null value.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
;                  $iType               - [optional] an integer value (1-5). Default is Null. The type of content to fill the control with. See Constants $LOW_FORM_CON_SOURCE_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sListContent        - [optional] a string value. Default is Null. Default is Null. The SQL statement, Table Name, etc., depending on the value of $iType.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComboBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oComboBox not a Combo Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bEmptyIsNull not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bInputRequired not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iType not an Integer, less than 1 or greater than 5. See Constants $LOW_FORM_CON_SOURCE_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $sListContent not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bEmptyIsNull
;                  |                               4 = Error setting $bInputRequired
;                  |                               8 = Error setting $iType
;                  |                               16 = Error setting $sListContent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConComboBoxGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConComboBoxData(ByRef $oComboBox, $sDataField = Null, $bEmptyIsNull = Null, $bInputRequired = Null, $iType = Null, $sListContent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[5]

	If Not IsObj($oComboBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oComboBox) <> $LOW_FORM_CON_TYPE_COMBO_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bEmptyIsNull, $bInputRequired, $iType, $sListContent) Then
		__LO_ArrayFill($avControl, $oComboBox.DataField(), $oComboBox.ConvertEmptyToNull(), $oComboBox.InputRequired(), _
				$oComboBox.ListSourceType(), $oComboBox.ListSource())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oComboBox.DataField = $sDataField
		$iError = ($oComboBox.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bEmptyIsNull <> Null) Then
		If Not IsBool($bEmptyIsNull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oComboBox.ConvertEmptyToNull = $bEmptyIsNull
		$iError = ($oComboBox.ConvertEmptyToNull() = $bEmptyIsNull) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oComboBox.InputRequired = $bInputRequired
		$iError = ($oComboBox.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iType <> Null) Then
		If Not __LO_IntIsBetween($iType, $LOW_FORM_CON_SOURCE_TYPE_TABLE, $LOW_FORM_CON_SOURCE_TYPE_TABLE_FIELDS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oComboBox.ListSourceType = $iType
		$iError = ($oComboBox.ListSourceType() = $iType) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sListContent <> Null) Then
		If Not IsString($sListContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oComboBox.ListSource = $sListContent
		$iError = ($oComboBox.ListSource() = $sListContent) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConComboBoxData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConComboBoxGeneral
; Description ...: Set or Retrieve general Table Control Combo Box Properties.
; Syntax ........: _LOWriter_FormConTableConComboBoxGeneral(ByRef $oComboBox[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $iMaxLen = Null[, $bEnabled = Null[, $bReadOnly = Null[, $iMouseScroll = Null[, $iWidth = Null[, $asList = Null[, $sDefaultTxt = Null[, $iAlign = Null[, $iLines = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]])
; Parameters ....: $oComboBox           - [in/out] an object. A Table Control Combo Box Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iMaxLen             - [optional] an integer value (-1-2147483647). Default is Null. The maximum text length that the Combo box will accept.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iWidth              - [optional] an integer value (100-200000). Default is Null. The width of the Column tab, in Hundredths of a Millimeter (HMM).
;                  $asList              - [optional] an array of strings. Default is Null. An array of entries. See remarks.
;                  $sDefaultTxt         - [optional] a string value. Default is Null. The default text of the combo Box.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLines              - [optional] an integer value. Default is Null. How many lines are shown in the dropdown list.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oComboBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oComboBox not a Combo Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iMaxLen not an Integer, less than -1 or greater than 2147483647.
;                  @Error 1 @Extended 7 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $iWidth not an Integer, less than 100 or greater than 200,000.
;                  @Error 1 @Extended 11 Return 0 = $asList not an Array.
;                  @Error 1 @Extended 12 Return ? = Element contained in $asList not a String. Returning problem element position.
;                  @Error 1 @Extended 13 Return 0 = $sDefaultTxt not a String.
;                  @Error 1 @Extended 14 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 15 Return 0 = $iLines not an Integer, less than -2147483648 or greater than 2147483647.
;                  @Error 1 @Extended 16 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 17 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 18 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 19 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $iMaxLen
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bReadOnly
;                  |                               64 = Error setting $iMouseScroll
;                  |                               128 = Error setting $iWidth
;                  |                               256 = Error setting $asList
;                  |                               512 = Error setting $sDefaultTxt
;                  |                               1024 = Error setting $iAlign
;                  |                               2048 = Error setting $iLines
;                  |                               4096 = Error setting $bHideSel
;                  |                               8192 = Error setting $sAddInfo
;                  |                               16384 = Error setting $sHelpText
;                  |                               32768 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 16 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $asList, $sDefaultTxt, $sAddInfo.
; Related .......: _LOWriter_FormConComboBoxData, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConComboBoxGeneral(ByRef $oComboBox, $sName = Null, $sLabel = Null, $iTxtDir = Null, $iMaxLen = Null, $bEnabled = Null, $bReadOnly = Null, $iMouseScroll = Null, $iWidth = Null, $asList = Null, $sDefaultTxt = Null, $iAlign = Null, $iLines = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[16]

	If Not IsObj($oComboBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oComboBox) <> $LOW_FORM_CON_TYPE_COMBO_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $iMaxLen, $bEnabled, $bReadOnly, $iMouseScroll, $iWidth, $asList, $sDefaultTxt, $iAlign, $iLines, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oComboBox.Name(), $oComboBox.Label(), $oComboBox.WritingMode(), $oComboBox.MaxTextLen(), _
				$oComboBox.Enabled(), $oComboBox.ReadOnly(), $oComboBox.MouseWheelBehavior(), _
				Int($oComboBox.Width() * 10), _ ; Multiply width by 10 to get Hundredths of a Millimeter value.
				$oComboBox.StringItemList(), $oComboBox.DefaultText(), _
				$oComboBox.Align(), $oComboBox.LineCount(), $oComboBox.HideInactiveSelection(), _
				$oComboBox.Tag(), $oComboBox.HelpText(), _
				$oComboBox.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oComboBox.Name = $sName
		$iError = ($oComboBox.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oComboBox.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oComboBox.Label = $sLabel
		$iError = ($oComboBox.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oComboBox.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oComboBox.WritingMode = $iTxtDir
		$iError = ($oComboBox.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iMaxLen = Default) Then
		$oComboBox.setPropertyToDefault("MaxTextLen")

	ElseIf ($iMaxLen <> Null) Then
		If Not __LO_IntIsBetween($iMaxLen, -1, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oComboBox.MaxTextLen = $iMaxLen
		$iError = ($oComboBox.MaxTextLen() = $iMaxLen) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oComboBox.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oComboBox.Enabled = $bEnabled
		$iError = ($oComboBox.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bReadOnly = Default) Then
		$oComboBox.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oComboBox.ReadOnly = $bReadOnly
		$iError = ($oComboBox.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iMouseScroll = Default) Then
		$oComboBox.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oComboBox.MouseWheelBehavior = $iMouseScroll
		$iError = ($oComboBox.MouseWheelBehavior = $iMouseScroll) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iWidth = Default) Then
		$oComboBox.setPropertyToDefault("Width")

	ElseIf ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 100, 200000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oComboBox.Width = Round($iWidth / 10) ; Divide Hundredths of a Millimeter value by 10 to obtain 10th MM.
		$iError = ($oComboBox.Width() = Round($iWidth / 10)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($asList = Default) Then
		$iError = BitOR($iError, 256) ; Can't Default StringItemList.

	ElseIf ($asList <> Null) Then
		If Not IsArray($asList) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		For $i = 0 To UBound($asList) - 1
			If Not IsString($asList[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, $i)

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next
		$oComboBox.StringItemList = $asList
		$iError = (UBound($oComboBox.StringItemList()) = UBound($asList)) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($sDefaultTxt = Default) Then
		$iError = BitOR($iError, 512) ; Can't Default DefaultText.

	ElseIf ($sDefaultTxt <> Null) Then
		If Not IsString($sDefaultTxt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oComboBox.DefaultText = $sDefaultTxt
		$iError = ($oComboBox.DefaultText() = $sDefaultTxt) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iAlign = Default) Then
		$oComboBox.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oComboBox.Align = $iAlign
		$iError = ($oComboBox.Align() = $iAlign) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($iLines = Default) Then
		$oComboBox.setPropertyToDefault("LineCount")

	ElseIf ($iLines <> Null) Then
		If Not __LO_IntIsBetween($iLines, -2147483648, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oComboBox.LineCount = $iLines
		$iError = ($oComboBox.LineCount = $iLines) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($bHideSel = Default) Then
		$oComboBox.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oComboBox.HideInactiveSelection = $bHideSel
		$iError = ($oComboBox.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 8192) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oComboBox.Tag = $sAddInfo
		$iError = ($oComboBox.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($sHelpText = Default) Then
		$oComboBox.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oComboBox.HelpText = $sHelpText
		$iError = ($oComboBox.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($sHelpURL = Default) Then
		$oComboBox.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oComboBox.HelpURL = $sHelpURL
		$iError = ($oComboBox.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConComboBoxGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConCurrencyFieldData
; Description ...: Set or Retrieve Table Control Currency Field Data Properties.
; Syntax ........: _LOWriter_FormConTableConCurrencyFieldData(ByRef $oCurrencyField[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oCurrencyField      - [in/out] an object. A Table Control Currency Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCurrencyField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oCurrencyField not a Currency Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConCurrencyFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConCurrencyFieldData(ByRef $oCurrencyField, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oCurrencyField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oCurrencyField) <> $LOW_FORM_CON_TYPE_CURRENCY_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oCurrencyField.DataField(), $oCurrencyField.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCurrencyField.DataField = $sDataField
		$iError = ($oCurrencyField.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCurrencyField.InputRequired = $bInputRequired
		$iError = ($oCurrencyField.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConCurrencyFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConCurrencyFieldGeneral
; Description ...: Set or Retrieve general Table Control Currency Field properties.
; Syntax ........: _LOWriter_FormConTableConCurrencyFieldGeneral(ByRef $oCurrencyField[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $bStrict = Null[, $bEnabled = Null[, $bReadOnly = Null[, $iMouseScroll = Null[, $nMin = Null[, $nMax = Null[, $iIncr = Null[, $nDefault = Null[, $iDecimal = Null[, $bThousandsSep = Null[, $sCurrSymbol = Null[, $bPrefix = Null[, $bSpin = Null[, $bRepeat = Null[, $iDelay = Null[, $iWidth = Null[, $iAlign = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oCurrencyField      - [in/out] an object. A Table Control Currency Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bStrict             - [optional] a boolean value. Default is Null. If True, strict formatting is enabled.
;                  $bEnabled            - [optional] a boolean value. Default is Null.If True, the control is enabled.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $nMin                - [optional] a general number value. Default is Null. The minimum value the control can be set to.
;                  $nMax                - [optional] a general number value. Default is Null. The maximum value the control can be set to.
;                  $iIncr               - [optional] an integer value. Default is Null. The amount to Increase or Decrease the value by.
;                  $nDefault            - [optional] a general number value. Default is Null. The default value the control will be set to.
;                  $iDecimal            - [optional] an integer value (0-20). Default is Null. The amount of decimal accuracy.
;                  $bThousandsSep       - [optional] a boolean value. Default is Null. If True, a thousands separator will be added.
;                  $sCurrSymbol         - [optional] a string value. Default is Null. The symbol to use for currency.
;                  $bPrefix             - [optional] a boolean value. Default is Null. If True, the currency symbol is prefixed to the value.
;                  $bSpin               - [optional] a boolean value. Default is Null. If True, the field will act as a spin button.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $iWidth              - [optional] an integer value (100-200000). Default is Null. The width of the Column tab, in Hundredths of a Millimeter (HMM).
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCurrencyField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oCurrencyField not a Currency Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bStrict not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $nMin not a Number.
;                  @Error 1 @Extended 11 Return 0 = $nMax not a Number.
;                  @Error 1 @Extended 12 Return 0 = $iIncr not an Integer.
;                  @Error 1 @Extended 13 Return 0 = $nDefault not a Number.
;                  @Error 1 @Extended 14 Return 0 = $iDecimal not an Integer, less than 0 or greater than 20.
;                  @Error 1 @Extended 15 Return 0 = $bThousandsSep not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $sCurrSymbol not a String.
;                  @Error 1 @Extended 17 Return 0 = $bPrefix not a Boolean.
;                  @Error 1 @Extended 18 Return 0 = $bSpin not a Boolean.
;                  @Error 1 @Extended 19 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 20 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 21 Return 0 = $iWidth not an Integer, less than 100 or greater than 200,000.
;                  @Error 1 @Extended 22 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 23 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 24 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 25 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 26 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bStrict
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bReadOnly
;                  |                               64 = Error setting $iMouseScroll
;                  |                               128 = Error setting $nMin
;                  |                               256 = Error setting $nMax
;                  |                               512 = Error setting $iIncr
;                  |                               1024 = Error setting $nDefault
;                  |                               2048 = Error setting $iDecimal
;                  |                               4096 = Error setting $bThousandsSep
;                  |                               8192 = Error setting $sCurrSymbol
;                  |                               16384 = Error setting $bPrefix
;                  |                               32768 = Error setting $bSpin
;                  |                               65536 = Error setting $bRepeat
;                  |                               131072 = Error setting $iDelay
;                  |                               262144 = Error setting $iWidth
;                  |                               524288 = Error setting $iAlign
;                  |                               1048576 = Error setting $bHideSel
;                  |                               2097152 = Error setting $sAddInfo
;                  |                               4194304 = Error setting $sHelpText
;                  |                               8388608 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 24 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $sAddInfo.
; Related .......: _LOWriter_FormConCurrencyFieldData, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConCurrencyFieldGeneral(ByRef $oCurrencyField, $sName = Null, $sLabel = Null, $iTxtDir = Null, $bStrict = Null, $bEnabled = Null, $bReadOnly = Null, $iMouseScroll = Null, $nMin = Null, $nMax = Null, $iIncr = Null, $nDefault = Null, $iDecimal = Null, $bThousandsSep = Null, $sCurrSymbol = Null, $bPrefix = Null, $bSpin = Null, $bRepeat = Null, $iDelay = Null, $iWidth = Null, $iAlign = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[24]

	If Not IsObj($oCurrencyField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oCurrencyField) <> $LOW_FORM_CON_TYPE_CURRENCY_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $bStrict, $bEnabled, $bReadOnly, $iMouseScroll, $nMin, $nMax, $iIncr, $nDefault, $iDecimal, $bThousandsSep, $sCurrSymbol, $bPrefix, $bSpin, $bRepeat, $iDelay, $iWidth, $iAlign, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oCurrencyField.Name(), $oCurrencyField.Label(), $oCurrencyField.WritingMode(), _
				$oCurrencyField.StrictFormat(), $oCurrencyField.Enabled(), $oCurrencyField.ReadOnly(), _
				$oCurrencyField.MouseWheelBehavior(), _
				$oCurrencyField.ValueMin(), $oCurrencyField.ValueMax(), $oCurrencyField.ValueStep(), $oCurrencyField.DefaultValue(), _
				$oCurrencyField.DecimalAccuracy(), $oCurrencyField.ShowThousandsSeparator(), $oCurrencyField.CurrencySymbol(), _
				$oCurrencyField.PrependCurrencySymbol(), $oCurrencyField.Spin(), $oCurrencyField.Repeat(), $oCurrencyField.RepeatDelay(), _
				Int($oCurrencyField.Width() * 10), _ ; Multiply width by 10 to get Hundredths of a Millimeter value.
				$oCurrencyField.Align(), $oCurrencyField.HideInactiveSelection(), $oCurrencyField.Tag(), _
				$oCurrencyField.HelpText(), $oCurrencyField.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCurrencyField.Name = $sName
		$iError = ($oCurrencyField.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oCurrencyField.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCurrencyField.Label = $sLabel
		$iError = ($oCurrencyField.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oCurrencyField.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oCurrencyField.WritingMode = $iTxtDir
		$iError = ($oCurrencyField.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bStrict = Default) Then
		$oCurrencyField.setPropertyToDefault("StrictFormat")

	ElseIf ($bStrict <> Null) Then
		If Not IsBool($bStrict) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oCurrencyField.StrictFormat = $bStrict
		$iError = ($oCurrencyField.StrictFormat() = $bStrict) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oCurrencyField.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oCurrencyField.Enabled = $bEnabled
		$iError = ($oCurrencyField.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bReadOnly = Default) Then
		$oCurrencyField.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oCurrencyField.ReadOnly = $bReadOnly
		$iError = ($oCurrencyField.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iMouseScroll = Default) Then
		$oCurrencyField.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oCurrencyField.MouseWheelBehavior = $iMouseScroll
		$iError = ($oCurrencyField.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($nMax = Default) Then
		$oCurrencyField.setPropertyToDefault("ValueMin")

	ElseIf ($nMin <> Null) Then
		If Not IsNumber($nMin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oCurrencyField.ValueMin = $nMin
		$iError = ($oCurrencyField.ValueMin() = $nMin) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($nMax = Default) Then
		$oCurrencyField.setPropertyToDefault("ValueMax")

	ElseIf ($nMax <> Null) Then
		If Not IsNumber($nMax) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oCurrencyField.ValueMax = $nMax
		$iError = ($oCurrencyField.ValueMax() = $nMax) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iIncr = Default) Then
		$oCurrencyField.setPropertyToDefault("ValueStep")

	ElseIf ($iIncr <> Null) Then
		If Not IsInt($iIncr) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oCurrencyField.ValueStep = $iIncr
		$iError = ($oCurrencyField.ValueStep() = $iIncr) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($nDefault = Default) Then
		$oCurrencyField.setPropertyToDefault("DefaultValue")

	ElseIf ($nDefault <> Null) Then
		If Not IsNumber($nDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oCurrencyField.DefaultValue = $nDefault
		$iError = ($oCurrencyField.DefaultValue() = $nDefault) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($iDecimal = Default) Then
		$oCurrencyField.setPropertyToDefault("DecimalAccuracy")

	ElseIf ($iDecimal <> Null) Then
		If Not __LO_IntIsBetween($iDecimal, 0, 20) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oCurrencyField.DecimalAccuracy = $iDecimal
		$iError = ($oCurrencyField.DecimalAccuracy() = $iDecimal) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($bThousandsSep = Default) Then
		$oCurrencyField.setPropertyToDefault("ShowThousandsSeparator")

	ElseIf ($bThousandsSep <> Null) Then
		If Not IsBool($bThousandsSep) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oCurrencyField.ShowThousandsSeparator = $bThousandsSep
		$iError = ($oCurrencyField.ShowThousandsSeparator() = $bThousandsSep) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($sCurrSymbol = Default) Then
		$oCurrencyField.setPropertyToDefault("CurrencySymbol")

	ElseIf ($sCurrSymbol <> Null) Then
		If Not IsString($sCurrSymbol) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oCurrencyField.CurrencySymbol = $sCurrSymbol
		$iError = ($oCurrencyField.CurrencySymbol() = $sCurrSymbol) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($bPrefix = Default) Then
		$oCurrencyField.setPropertyToDefault("PrependCurrencySymbol")

	ElseIf ($bPrefix <> Null) Then
		If Not IsBool($bPrefix) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oCurrencyField.PrependCurrencySymbol = $bPrefix
		$iError = ($oCurrencyField.PrependCurrencySymbol() = $bPrefix) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($bSpin = Default) Then
		$oCurrencyField.setPropertyToDefault("Spin")

	ElseIf ($bSpin <> Null) Then
		If Not IsBool($bSpin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oCurrencyField.Spin = $bSpin
		$iError = ($oCurrencyField.Spin() = $bSpin) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bRepeat = Default) Then
		$oCurrencyField.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oCurrencyField.Repeat = $bRepeat
		$iError = ($oCurrencyField.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($iDelay = Default) Then
		$oCurrencyField.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oCurrencyField.RepeatDelay = $iDelay
		$iError = ($oCurrencyField.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($iWidth = Default) Then
		$oCurrencyField.setPropertyToDefault("Width")

	ElseIf ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 100, 200000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oCurrencyField.Width = Round($iWidth / 10) ; Divide Hundredths of a Millimeter value by 10 to obtain 10th MM.
		$iError = ($oCurrencyField.Width() = Round($iWidth / 10)) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($iAlign = Default) Then
		$oCurrencyField.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oCurrencyField.Align = $iAlign
		$iError = ($oCurrencyField.Align() = $iAlign) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($bHideSel = Default) Then
		$oCurrencyField.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oCurrencyField.HideInactiveSelection = $bHideSel
		$iError = ($oCurrencyField.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 2097152) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oCurrencyField.Tag = $sAddInfo
		$iError = ($oCurrencyField.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($sHelpText = Default) Then
		$oCurrencyField.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oCurrencyField.HelpText = $sHelpText
		$iError = ($oCurrencyField.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	If ($sHelpURL = Default) Then
		$oCurrencyField.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, 0)

		$oCurrencyField.HelpURL = $sHelpURL
		$iError = ($oCurrencyField.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 8388608))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConCurrencyFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConDateFieldData
; Description ...: Set or Retrieve Table Control Date Field Data Properties.
; Syntax ........: _LOWriter_FormConTableConDateFieldData(ByRef $oDateField[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oDateField          - [in/out] an object. A Table Control Date Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDateField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oDateField not a Date Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConDateFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConDateFieldData(ByRef $oDateField, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oDateField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oDateField) <> $LOW_FORM_CON_TYPE_DATE_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oDateField.DataField(), $oDateField.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDateField.DataField = $sDataField
		$iError = ($oDateField.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDateField.InputRequired = $bInputRequired
		$iError = ($oDateField.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConDateFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConDateFieldGeneral
; Description ...: Set or Retrieve general Table Control Date Field properties.
; Syntax ........: _LOWriter_FormConTableConDateFieldGeneral(ByRef $oDateField[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $bStrict = Null[, $bEnabled = Null[, $bReadOnly = Null[, $iMouseScroll = Null[, $tDateMin = Null[, $tDateMax = Null[, $iFormat = Null[, $tDateDefault = Null[, $bSpin = Null[, $bRepeat = Null[, $iDelay = Null[, $iWidth = Null[, $iAlign = Null[, $bDropdown = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oDateField          - [in/out] an object. A Table Control Date Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bStrict             - [optional] a boolean value. Default is Null. If True, strict formatting is enabled.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $tDateMin            - [optional] a dll struct value. Default is Null. The minimum date allowed to be entered, created previously by _LOWriter_DateStructCreate.
;                  $tDateMax            - [optional] a dll struct value. Default is Null. The maximum date allowed to be entered, created previously by _LOWriter_DateStructCreate.
;                  $iFormat             - [optional] an integer value (0-11). Default is Null. The Date Format to display the content in. See Constants $LOW_FORM_CON_DATE_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $tDateDefault        - [optional] a dll struct value. Default is Null. The Default date to display, created previously by _LOWriter_DateStructCreate.
;                  $bSpin               - [optional] a boolean value. Default is Null. If True, the field will act as a spin button.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $iWidth              - [optional] an integer value (100-200000). Default is Null. The width of the Column tab, in Hundredths of a Millimeter (HMM).
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bDropdown           - [optional] a boolean value. Default is Null. If True, the field will behave as a dropdown control.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDateField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oDateField not a Date Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bStrict not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $tDateMin not an Object.
;                  @Error 1 @Extended 11 Return 0 = $tDateMax not an Object.
;                  @Error 1 @Extended 12 Return 0 = $iFormat not an Integer, less then 0 or greater than 11. See Constants $LOW_FORM_CON_DATE_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $tDateDefault not an Object.
;                  @Error 1 @Extended 14 Return 0 = $bSpin not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 17 Return 0 = $iWidth not an Integer, less than 100 or greater than 200,000.
;                  @Error 1 @Extended 18 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 19 Return 0 = $bDropdown not a Boolean.
;                  @Error 1 @Extended 20 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 21 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 22 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 23 Return 0 = $sHelpURL not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.DateTime" Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.util.Date" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current Minimum Date.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve current Maximum Date.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bStrict
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bReadOnly
;                  |                               64 = Error setting $iMouseScroll
;                  |                               128 = Error setting $tDateMin
;                  |                               256 = Error setting $tDateMax
;                  |                               512 = Error setting $iFormat
;                  |                               1024 = Error setting $tDateDefault
;                  |                               2048 = Error setting $bSpin
;                  |                               4096 = Error setting $bRepeat
;                  |                               8192 = Error setting $iDelay
;                  |                               16384 = Error setting $iWidth
;                  |                               32768 = Error setting $iAlign
;                  |                               65536 = Error setting $bDropdown
;                  |                               131072 = Error setting $bHideSel
;                  |                               262144 = Error setting $sAddInfo
;                  |                               524288 = Error setting $sHelpText
;                  |                               1048576 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 21 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $sAddInfo.
; Related .......: _LOWriter_FormConDateFieldData, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConDateFieldGeneral(ByRef $oDateField, $sName = Null, $sLabel = Null, $iTxtDir = Null, $bStrict = Null, $bEnabled = Null, $bReadOnly = Null, $iMouseScroll = Null, $tDateMin = Null, $tDateMax = Null, $iFormat = Null, $tDateDefault = Null, $bSpin = Null, $bRepeat = Null, $iDelay = Null, $iWidth = Null, $iAlign = Null, $bDropdown = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tDate, $tCurMin, $tCurMax, $tCurDefault
	Local $avControl[21]

	If Not IsObj($oDateField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oDateField) <> $LOW_FORM_CON_TYPE_DATE_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $bStrict, $bEnabled, $bReadOnly, $iMouseScroll, $tDateMin, $tDateMax, $iFormat, $tDateDefault, $bSpin, $bRepeat, $iDelay, $iWidth, $iAlign, $bDropdown, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		$tDate = $oDateField.DateMin()
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$tCurMin = __LO_CreateStruct("com.sun.star.util.DateTime")
		If Not IsObj($tCurMin) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tCurMin.Year = $tDate.Year()
		$tCurMin.Month = $tDate.Month()
		$tCurMin.Day = $tDate.Day()

		$tDate = $oDateField.DateMax()
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$tCurMax = __LO_CreateStruct("com.sun.star.util.DateTime")
		If Not IsObj($tCurMax) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tCurMax.Year = $tDate.Year()
		$tCurMax.Month = $tDate.Month()
		$tCurMax.Day = $tDate.Day()

		$tDate = $oDateField.DefaultDate() ; Default date is Null when not set.
		If IsObj($tDate) Then
			$tCurDefault = __LO_CreateStruct("com.sun.star.util.DateTime")
			If Not IsObj($tCurDefault) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tCurDefault.Year = $tDate.Year()
			$tCurDefault.Month = $tDate.Month()
			$tCurDefault.Day = $tDate.Day()

		Else
			$tCurDefault = $tDate
		EndIf

		__LO_ArrayFill($avControl, $oDateField.Name(), $oDateField.Label(), $oDateField.WritingMode(), $oDateField.StrictFormat(), _
				$oDateField.Enabled(), $oDateField.ReadOnly(), $oDateField.MouseWheelBehavior(), _
				$tCurMin, $tCurMax, $oDateField.DateFormat(), $tCurDefault, $oDateField.Spin(), _
				$oDateField.Repeat(), $oDateField.RepeatDelay(), Int($oDateField.Width() * 10), _ ; Multiply width by 10 to get Hundredths of a Millimeter value.
				$oDateField.Align(), _
				$oDateField.Dropdown(), _
				$oDateField.HideInactiveSelection(), $oDateField.Tag(), $oDateField.HelpText(), $oDateField.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDateField.Name = $sName
		$iError = ($oDateField.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oDateField.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDateField.Label = $sLabel
		$iError = ($oDateField.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oDateField.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oDateField.WritingMode = $iTxtDir
		$iError = ($oDateField.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bStrict = Default) Then
		$oDateField.setPropertyToDefault("StrictFormat")

	ElseIf ($bStrict <> Null) Then
		If Not IsBool($bStrict) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oDateField.StrictFormat = $bStrict
		$iError = ($oDateField.StrictFormat() = $bStrict) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oDateField.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oDateField.Enabled = $bEnabled
		$iError = ($oDateField.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bReadOnly = Default) Then
		$oDateField.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oDateField.ReadOnly = $bReadOnly
		$iError = ($oDateField.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iMouseScroll = Default) Then
		$oDateField.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oDateField.MouseWheelBehavior = $iMouseScroll
		$iError = ($oDateField.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($tDateMin = Default) Then
		$oDateField.setPropertyToDefault("DateMin")

	ElseIf ($tDateMin <> Null) Then
		If Not IsObj($tDateMin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tDate = __LO_CreateStruct("com.sun.star.util.Date")
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tDate.Year = $tDateMin.Year()
		$tDate.Month = $tDateMin.Month()
		$tDate.Day = $tDateMin.Day()

		$oDateField.DateMin = $tDate
		$iError = (__LOWriter_DateStructCompare($oDateField.DateMin(), $tDate, True)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($tDateMax = Default) Then
		$oDateField.setPropertyToDefault("DateMax")

	ElseIf ($tDateMax <> Null) Then
		If Not IsObj($tDateMax) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tDate = __LO_CreateStruct("com.sun.star.util.Date")
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tDate.Year = $tDateMax.Year()
		$tDate.Month = $tDateMax.Month()
		$tDate.Day = $tDateMax.Day()

		$oDateField.DateMax = $tDate
		$iError = (__LOWriter_DateStructCompare($oDateField.DateMax(), $tDate, True)) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iFormat = Default) Then
		$oDateField.setPropertyToDefault("DateFormat")

	ElseIf ($iFormat <> Null) Then
		If Not __LO_IntIsBetween($iFormat, $LOW_FORM_CON_DATE_FRMT_SYSTEM_SHORT, $LOW_FORM_CON_DATE_FRMT_SHORT_YYYY_MM_DD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oDateField.DateFormat = $iFormat
		$iError = ($oDateField.DateFormat() = $iFormat) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($tDateDefault = Default) Then
		$oDateField.setPropertyToDefault("DefaultDate")

	ElseIf ($tDateDefault <> Null) Then
		If Not IsObj($tDateDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$tDate = __LO_CreateStruct("com.sun.star.util.Date")
		If Not IsObj($tDate) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tDate.Year = $tDateDefault.Year()
		$tDate.Month = $tDateDefault.Month()
		$tDate.Day = $tDateDefault.Day()

		$oDateField.DefaultDate = $tDate
		$iError = (__LOWriter_DateStructCompare($oDateField.DefaultDate(), $tDate, True)) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($bSpin = Default) Then
		$oDateField.setPropertyToDefault("Spin")

	ElseIf ($bSpin <> Null) Then
		If Not IsBool($bSpin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oDateField.Spin = $bSpin
		$iError = ($oDateField.Spin() = $bSpin) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($bRepeat = Default) Then
		$oDateField.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oDateField.Repeat = $bRepeat
		$iError = ($oDateField.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iDelay = Default) Then
		$oDateField.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oDateField.RepeatDelay = $iDelay
		$iError = ($oDateField.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iWidth = Default) Then
		$oDateField.setPropertyToDefault("Width")

	ElseIf ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 100, 200000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oDateField.Width = Round($iWidth / 10) ; Divide Hundredths of a Millimeter value by 10 to obtain 10th MM.
		$iError = ($oDateField.Width() = Round($iWidth / 10)) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iAlign = Default) Then
		$oDateField.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oDateField.Align = $iAlign
		$iError = ($oDateField.Align() = $iAlign) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bDropdown = Default) Then
		$oDateField.setPropertyToDefault("Dropdown")

	ElseIf ($bDropdown <> Null) Then
		If Not IsBool($bDropdown) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oDateField.Dropdown = $bDropdown
		$iError = ($oDateField.Dropdown() = $bDropdown) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($bHideSel = Default) Then
		$oDateField.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oDateField.HideInactiveSelection = $bHideSel
		$iError = ($oDateField.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 262144) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oDateField.Tag = $sAddInfo
		$iError = ($oDateField.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($sHelpText = Default) Then
		$oDateField.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oDateField.HelpText = $sHelpText
		$iError = ($oDateField.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($sHelpURL = Default) Then
		$oDateField.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oDateField.HelpURL = $sHelpURL
		$iError = ($oDateField.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConDateFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConFormattedFieldData
; Description ...: Set or Retrieve Table Control Formatted Field Data Properties.
; Syntax ........: _LOWriter_FormConTableConFormattedFieldData(ByRef $oFormatField[, $sDataField = Null[, $bEmptyIsNull = Null[, $bInputRequired = Null[, $bFilter = Null]]]])
; Parameters ....: $oFormatField        - [in/out] an object. A Table Control Formatted Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bEmptyIsNull        - [optional] a boolean value. Default is Null. If True, an empty string will be treated as a Null value.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
;                  $bFilter             - [optional] a boolean value. Default is Null. If True, filter proposal is active.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormatField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oFormatField not a Formatted Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bEmptyIsNull not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bInputRequired not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bFilter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bEmptyIsNull
;                  |                               4 = Error setting $bInputRequired
;                  |                               8 = Error setting $bFilter
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConFormattedFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConFormattedFieldData(ByRef $oFormatField, $sDataField = Null, $bEmptyIsNull = Null, $bInputRequired = Null, $bFilter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[4]

	If Not IsObj($oFormatField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oFormatField) <> $LOW_FORM_CON_TYPE_FORMATTED_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bEmptyIsNull, $bInputRequired, $bFilter) Then
		__LO_ArrayFill($avControl, $oFormatField.DataField(), $oFormatField.ConvertEmptyToNull(), $oFormatField.InputRequired(), _
				$oFormatField.UseFilterValueProposal())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFormatField.DataField = $sDataField
		$iError = ($oFormatField.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bEmptyIsNull <> Null) Then
		If Not IsBool($bEmptyIsNull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFormatField.ConvertEmptyToNull = $bEmptyIsNull
		$iError = ($oFormatField.ConvertEmptyToNull() = $bEmptyIsNull) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFormatField.InputRequired = $bInputRequired
		$iError = ($oFormatField.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bFilter <> Null) Then
		If Not IsBool($bFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFormatField.UseFilterValueProposal = $bFilter
		$iError = ($oFormatField.UseFilterValueProposal() = $bFilter) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConFormattedFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConFormattedFieldGeneral
; Description ...: Set or Retrieve general Table Control Formatted Field properties.
; Syntax ........: _LOWriter_FormConTableConFormattedFieldGeneral(ByRef $oFormatField[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $iMaxLen = Null[, $bEnabled = Null[, $bReadOnly = Null[, $iMouseScroll = Null[, $nMin = Null[, $nMax = Null[, $nDefault = Null[, $iFormat = Null[, $bSpin = Null[, $bRepeat = Null[, $iDelay = Null[, $iWidth = Null[, $iAlign = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oFormatField        - [in/out] an object. A Table Control Formatted Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iMaxLen             - [optional] an integer value (-1-2147483647). Default is Null. The maximum text length that the Formatted field will accept.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $nMin                - [optional] a general number value. Default is Null. The minimum value allowed in the field.
;                  $nMax                - [optional] a general number value. Default is Null. The maximum value allowed in the field.
;                  $nDefault            - [optional] a general number value. Default is Null. The default value of the field.
;                  $iFormat             - [optional] an integer value. Default is Null. The Number Format Key to display the content in, retrieved from a previous _LOWriter_FormatKeysGetList call, or created by _LOWriter_FormatKeyCreate function.
;                  $bSpin               - [optional] a boolean value. Default is Null. If True, the field will act as a spin button.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $iWidth              - [optional] an integer value (100-200000). Default is Null. The width of the Column tab, in Hundredths of a Millimeter (HMM).
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormatField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oFormatField not a Formatted Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iMaxLen not an Integer, less than -1 or greater than 2147483647.
;                  @Error 1 @Extended 7 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $nMin not a Number.
;                  @Error 1 @Extended 11 Return 0 = $nMax not a Number.
;                  @Error 1 @Extended 12 Return 0 = $nDefault not a Number.
;                  @Error 1 @Extended 13 Return 0 = $iFormat not an Integer.
;                  @Error 1 @Extended 14 Return 0 = Format key called in $iFormat not found in document.
;                  @Error 1 @Extended 15 Return 0 = $bSpin not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 17 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 18 Return 0 = $iWidth not an Integer, less than 100 or greater than 200,000.
;                  @Error 1 @Extended 19 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 20 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 21 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 22 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 23 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $iMaxLen
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bReadOnly
;                  |                               64 = Error setting $iMouseScroll
;                  |                               128 = Error setting $nMin
;                  |                               256 = Error setting $nMax
;                  |                               512 = Error setting $nDefault
;                  |                               1024 = Error setting $iFormat
;                  |                               2048 = Error setting $bSpin
;                  |                               4096 = Error setting $bRepeat
;                  |                               8192 = Error setting $iDelay
;                  |                               16384 = Error setting $iWidth
;                  |                               32768 = Error setting $iAlign
;                  |                               65536 = Error setting $bHideSel
;                  |                               131072 = Error setting $sAddInfo
;                  |                               262144 = Error setting $sHelpText
;                  |                               524288 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 20 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $sAddInfo.
; Related .......: _LOWriter_FormatKeyCreate, _LOWriter_FormatKeysGetList, _LOWriter_FormConFormattedFieldData, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConFormattedFieldGeneral(ByRef $oFormatField, $sName = Null, $sLabel = Null, $iTxtDir = Null, $iMaxLen = Null, $bEnabled = Null, $bReadOnly = Null, $iMouseScroll = Null, $nMin = Null, $nMax = Null, $nDefault = Null, $iFormat = Null, $bSpin = Null, $bRepeat = Null, $iDelay = Null, $iWidth = Null, $iAlign = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oDoc
	Local $avControl[20]

	If Not IsObj($oFormatField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oFormatField) <> $LOW_FORM_CON_TYPE_FORMATTED_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $iMaxLen, $bEnabled, $bReadOnly, $iMouseScroll, $nMin, $nMax, $nDefault, $iFormat, $bSpin, $bRepeat, $iDelay, $iWidth, $iAlign, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oFormatField.Name(), $oFormatField.Label(), $oFormatField.WritingMode(), $oFormatField.MaxTextLen(), _
				$oFormatField.Enabled(), $oFormatField.ReadOnly(), _
				$oFormatField.MouseWheelBehavior(), $oFormatField.EffectiveMin(), _
				$oFormatField.EffectiveMax(), $oFormatField.EffectiveDefault(), $oFormatField.FormatKey(), $oFormatField.Spin(), _
				$oFormatField.Repeat(), $oFormatField.RepeatDelay(), Int($oFormatField.Width() * 10), _ ; Multiply width by 10 to get Hundredths of a Millimeter value.
				$oFormatField.Align(), _
				$oFormatField.HideInactiveSelection(), $oFormatField.Tag(), $oFormatField.HelpText(), $oFormatField.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFormatField.Name = $sName
		$iError = ($oFormatField.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oFormatField.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFormatField.Label = $sLabel
		$iError = ($oFormatField.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oFormatField.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFormatField.WritingMode = $iTxtDir
		$iError = ($oFormatField.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iMaxLen = Default) Then
		$oFormatField.setPropertyToDefault("MaxTextLen")

	ElseIf ($iMaxLen <> Null) Then
		If Not __LO_IntIsBetween($iMaxLen, -1, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFormatField.MaxTextLen = $iMaxLen
		$iError = ($oFormatField.MaxTextLen = $iMaxLen) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oFormatField.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFormatField.Enabled = $bEnabled
		$iError = ($oFormatField.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bReadOnly = Default) Then
		$oFormatField.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFormatField.ReadOnly = $bReadOnly
		$iError = ($oFormatField.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iMouseScroll = Default) Then
		$oFormatField.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oFormatField.MouseWheelBehavior = $iMouseScroll
		$iError = ($oFormatField.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($nMin = Default) Then
		$oFormatField.setPropertyToDefault("EffectiveMin")

	ElseIf ($nMin <> Null) Then
		If Not IsNumber($nMin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oFormatField.EffectiveMin = $nMin
		$iError = ($oFormatField.EffectiveMin() = $nMin) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($nMax = Default) Then
		$oFormatField.setPropertyToDefault("EffectiveMax")

	ElseIf ($nMax <> Null) Then
		If Not IsNumber($nMax) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oFormatField.EffectiveMax = $nMax
		$iError = ($oFormatField.EffectiveMax() = $nMax) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($nDefault = Default) Then
		$oFormatField.setPropertyToDefault("EffectiveDefault")

	ElseIf ($nDefault <> Null) Then
		If Not IsNumber($nDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oFormatField.EffectiveDefault = $nDefault
		$iError = ($oFormatField.EffectiveDefault() = $nDefault) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iFormat = Default) Then
		$oFormatField.setPropertyToDefault("FormatKey")

	ElseIf ($iFormat <> Null) Then
		If Not IsInt($iFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oDoc = $oFormatField.Parent() ; Identify the parent document.
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Do
			$oDoc = $oDoc.getParent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Until $oDoc.supportsService("com.sun.star.text.TextDocument")
		If Not _LOWriter_FormatKeyExists($oDoc, $iFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oFormatField.FormatKey = $iFormat
		$iError = ($oFormatField.FormatKey() = $iFormat) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($bSpin = Default) Then
		$oFormatField.setPropertyToDefault("Spin")

	ElseIf ($bSpin <> Null) Then
		If Not IsBool($bSpin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oFormatField.Spin = $bSpin
		$iError = ($oFormatField.Spin() = $bSpin) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($bRepeat = Default) Then
		$oFormatField.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oFormatField.Repeat = $bRepeat
		$iError = ($oFormatField.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iDelay = Default) Then
		$oFormatField.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oFormatField.RepeatDelay = $iDelay
		$iError = ($oFormatField.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iWidth = Default) Then
		$oFormatField.setPropertyToDefault("Width")

	ElseIf ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 100, 200000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oFormatField.Width = Round($iWidth / 10) ; Divide Hundredths of a Millimeter value by 10 to obtain 10th MM.
		$iError = ($oFormatField.Width() = Round($iWidth / 10)) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iAlign = Default) Then
		$oFormatField.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oFormatField.Align = $iAlign
		$iError = ($oFormatField.Align() = $iAlign) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bHideSel = Default) Then
		$oFormatField.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oFormatField.HideInactiveSelection = $bHideSel
		$iError = ($oFormatField.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 131072) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oFormatField.Tag = $sAddInfo
		$iError = ($oFormatField.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($sHelpText = Default) Then
		$oFormatField.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oFormatField.HelpText = $sHelpText
		$iError = ($oFormatField.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($sHelpURL = Default) Then
		$oFormatField.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oFormatField.HelpURL = $sHelpURL
		$iError = ($oFormatField.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConFormattedFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConGeneral
; Description ...: Set or Retrieve general Table Control properties.
; Syntax ........: _LOWriter_FormConTableConGeneral(ByRef $oTableCon[, $sName = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bVisible = Null[, $bPrintable = Null[, $bTabStop = Null[, $iTabOrder = Null[, $mFont = Null[, $nRowHeight = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bNavBar = Null[, $bRecordMarker = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]])
; Parameters ....: $oTableCon           - [in/out] an object. A Table Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $nRowHeight          - [optional] a general number value (-21474836.48-21474836.48). Default is Null. The Row height, set in Hundredths of a Millimeter (HMM).
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bNavBar             - [optional] a boolean value. Default is Null. If True, the Navigation Bar is displayed on the lower border of the Table control.
;                  $bRecordMarker       - [optional] a boolean value. Default is Null. If True, a record marker is displayed.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTableCon not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTableCon not a Table Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 10 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 11 Return 0 = $nRowHeight not a Number, less than -21474836.48 or greater than 21474836.48.
;                  @Error 1 @Extended 12 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 13 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 14 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 15 Return 0 = $bNavBar not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $bRecordMarker not a Boolean.
;                  @Error 1 @Extended 17 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 18 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 19 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $iTxtDir
;                  |                               4 = Error setting $bEnabled
;                  |                               8 = Error setting $bVisible
;                  |                               16 = Error setting $bPrintable
;                  |                               32 = Error setting $bTabStop
;                  |                               64 = Error setting $iTabOrder
;                  |                               128 = Error setting $mFont
;                  |                               256 = Error setting $nRowHeight
;                  |                               512 = Error setting $iBackColor
;                  |                               1024 = Error setting $iBorder
;                  |                               2048 = Error setting $iBorderColor
;                  |                               4096 = Error setting $bNavBar
;                  |                               8192 = Error setting $bRecordMarker
;                  |                               16384 = Error setting $sAddInfo
;                  |                               32768 = Error setting $sHelpText
;                  |                               65536 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 17 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConGeneral(ByRef $oTableCon, $sName = Null, $iTxtDir = Null, $bEnabled = Null, $bVisible = Null, $bPrintable = Null, $bTabStop = Null, $iTabOrder = Null, $mFont = Null, $nRowHeight = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bNavBar = Null, $bRecordMarker = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[17]

	If Not IsObj($oTableCon) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTableCon) <> $LOW_FORM_CON_TYPE_TABLE_CONTROL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $iTxtDir, $bEnabled, $bVisible, $bPrintable, $bTabStop, $iTabOrder, $mFont, $nRowHeight, $iBackColor, $iBorder, $iBorderColor, $bNavBar, $bRecordMarker, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oTableCon.Control.Name(), $oTableCon.Control.WritingMode(), $oTableCon.Control.Enabled(), $oTableCon.Control.EnableVisible(), _
				$oTableCon.Control.Printable(), $oTableCon.Control.Tabstop(), $oTableCon.Control.TabIndex(), __LOWriter_FormConSetGetFontDesc($oTableCon), $oTableCon.Control.RowHeight(), _
				$oTableCon.Control.BackgroundColor(), $oTableCon.Control.Border(), $oTableCon.Control.BorderColor(), $oTableCon.Control.HasNavigationBar(), _
				$oTableCon.Control.HasRecordMarker(), $oTableCon.Control.Tag(), $oTableCon.Control.HelpText(), $oTableCon.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTableCon.Control.Name = $sName
		$iError = ($oTableCon.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTxtDir = Default) Then
		$oTableCon.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTableCon.Control.WritingMode = $iTxtDir
		$iError = ($oTableCon.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bEnabled = Default) Then
		$oTableCon.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTableCon.Control.Enabled = $bEnabled
		$iError = ($oTableCon.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bVisible = Default) Then
		$oTableCon.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTableCon.Control.EnableVisible = $bVisible
		$iError = ($oTableCon.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bPrintable = Default) Then
		$oTableCon.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oTableCon.Control.Printable = $bPrintable
		$iError = ($oTableCon.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bTabStop = Default) Then
		$oTableCon.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oTableCon.Control.Tabstop = $bTabStop
		$iError = ($oTableCon.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 64) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oTableCon.Control.TabIndex = $iTabOrder
		$iError = ($oTableCon.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 128) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		__LOWriter_FormConSetGetFontDesc($oTableCon, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($nRowHeight = Default) Then
		$oTableCon.Control.setPropertyToDefault("RowHeight")

	ElseIf ($nRowHeight <> Null) Then
		If Not __LO_NumIsBetween($nRowHeight, -21474836.48, 21474836.48) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oTableCon.Control.RowHeight = $nRowHeight
		$iError = (__LO_NumIsBetween($oTableCon.Control.RowHeight(), $nRowHeight - 1, $nRowHeight + 1)) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iBackColor = Default) Then
		$oTableCon.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oTableCon.Control.BackgroundColor = $iBackColor
		$iError = ($oTableCon.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iBorder = Default) Then
		$oTableCon.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oTableCon.Control.Border = $iBorder
		$iError = ($oTableCon.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($iBorderColor = Default) Then
		$oTableCon.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oTableCon.Control.setPropertyValue("BorderColor", $iBorderColor)
		$iError = ($oTableCon.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($bNavBar = Default) Then
		$oTableCon.Control.setPropertyToDefault("HasNavigationBar")

	ElseIf ($bNavBar <> Null) Then
		If Not IsBool($bNavBar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oTableCon.Control.HasNavigationBar = $bNavBar
		$iError = ($oTableCon.Control.HasNavigationBar() = $bNavBar) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($bRecordMarker = Default) Then
		$oTableCon.Control.setPropertyToDefault("HasRecordMarker")

	ElseIf ($bRecordMarker <> Null) Then
		If Not IsBool($bRecordMarker) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oTableCon.Control.HasRecordMarker = $bRecordMarker
		$iError = ($oTableCon.Control.HasRecordMarker() = $bRecordMarker) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 8192) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oTableCon.Control.Tag = $sAddInfo
		$iError = ($oTableCon.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($sHelpText = Default) Then
		$oTableCon.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oTableCon.Control.HelpText = $sHelpText
		$iError = ($oTableCon.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($sHelpURL = Default) Then
		$oTableCon.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oTableCon.Control.HelpURL = $sHelpURL
		$iError = ($oTableCon.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConListBoxData
; Description ...: Set or Retrieve Table Control List Box Data Properties.
; Syntax ........: _LOWriter_FormConTableConListBoxData(ByRef $oListBox[, $sDataField = Null[, $bInputRequired = Null[, $iType = Null[, $asListContent = Null[, $iBoundField = Null]]]]])
; Parameters ....: $oListBox            - [in/out] an object. A Table Control List Box Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
;                  $iType               - [optional] an integer value (0-5). Default is Null. The type of content to fill the control with. See Constants $LOW_FORM_CON_SOURCE_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $asListContent       - [optional] an array of strings. Default is Null. A single dimension array. See remarks
;                  $iBoundField         - [optional] an integer value (-1-2147483647). Default is Null. The bound data field of a linked table to display.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oListBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oListBox not a List Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iType not an Integer, less than 0 or greater than 5. See Constants $LOW_FORM_CON_SOURCE_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $asListContent not an Array.
;                  @Error 1 @Extended 7 Return 0 = $iType not set to Valuelist and array called in $asListContent has more than 1 element.
;                  @Error 1 @Extended 8 Return 0 = $iBoundField not an Integer, less than -1 or greater than 2147483647.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  |                               4 = Error setting $iType
;                  |                               8 = Error setting $asListContent
;                  |                               16 = Error setting $iBoundField
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
;                  $asListContent is not error checked for the same content, but only that the set array size is the same.
;                  $asListContent should be a single dimension array with a appropriate value in each element. e.g. If $iType is set to Table, the element will contain a Table name. Or if $iType is set to Value List, each element will contain a list item.
;                  For types other than Value list for $iType, the array sound contain a single element.
; Related .......: _LOWriter_FormConListBoxGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConListBoxData(ByRef $oListBox, $sDataField = Null, $bInputRequired = Null, $iType = Null, $asListContent = Null, $iBoundField = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[5]

	If Not IsObj($oListBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oListBox) <> $LOW_FORM_CON_TYPE_LIST_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired, $iType, $asListContent, $iBoundField) Then
		__LO_ArrayFill($avControl, $oListBox.DataField(), $oListBox.InputRequired(), $oListBox.ListSourceType(), _
				$oListBox.ListSource(), $oListBox.BoundColumn())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oListBox.DataField = $sDataField
		$iError = ($oListBox.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oListBox.InputRequired = $bInputRequired
		$iError = ($oListBox.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iType <> Null) Then
		If Not __LO_IntIsBetween($iType, $LOW_FORM_CON_SOURCE_TYPE_VALUE_LIST, $LOW_FORM_CON_SOURCE_TYPE_TABLE_FIELDS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oListBox.ListSourceType = $iType
		$iError = ($oListBox.ListSourceType() = $iType) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($asListContent <> Null) Then
		If Not IsArray($asListContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If ($oListBox.ListSourceType() <> $LOW_FORM_CON_SOURCE_TYPE_VALUE_LIST) And (UBound($asListContent) > 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oListBox.ListSource = $asListContent
		$iError = (UBound($oListBox.ListSource()) = UBound($asListContent)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iBoundField <> Null) Then
		If Not __LO_IntIsBetween($iBoundField, -1, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oListBox.BoundColumn = $iBoundField
		$iError = ($oListBox.BoundColumn() = $iBoundField) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConListBoxData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConListBoxGeneral
; Description ...: Set or Retrieve general Table Control List box properties.
; Syntax ........: _LOWriter_FormConTableConListBoxGeneral(ByRef $oListBox[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $bEnabled = Null[, $bReadOnly = Null[, $iMouseScroll = Null[, $iWidth = Null[, $asList = Null[, $iAlign = Null[, $iLines = Null[, $aiDefaultSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]])
; Parameters ....: $oListBox            - [in/out] an object. A Table Control List Box Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iWidth              - [optional] an integer value (100-200000). Default is Null. The width of the Column tab, in Hundredths of a Millimeter (HMM).
;                  $asList              - [optional] an array of strings. Default is Null. An array of List entries. See remarks.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLines              - [optional] an integer value (-2147483648-2147483647). Default is Null. If $bDropdown is True, $iLines specifies how many lines are shown in the dropdown list.
;                  $aiDefaultSel        - [optional] an array of integers. Default is Null. A single dimension array of selection values. See remarks.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oListBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oListBox not a List Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $iWidth not an Integer, less than 100 or greater than 200,000.
;                  @Error 1 @Extended 10 Return 0 = $asList not an Array.
;                  @Error 1 @Extended 11 Return ? = Element contained in $asList not a String. Returning problem element position.
;                  @Error 1 @Extended 12 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $iLines not an Integer, less than -2147483648 or greater than 2147483647.
;                  @Error 1 @Extended 14 Return 0 = $aiDefaultSel not an Array.
;                  @Error 1 @Extended 15 Return ? = Element contained in $aiDefaultSel not an Integer. Returning problem element position.
;                  @Error 1 @Extended 16 Return ? = Integer contained in Element of $aiDefaultSel greater than number of List items. Returning problem element position.
;                  @Error 1 @Extended 17 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 18 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 19 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bEnabled
;                  |                               16 = Error setting $bReadOnly
;                  |                               32 = Error setting $iMouseScroll
;                  |                               64 = Error setting $iWidth
;                  |                               128 = Error setting $asList
;                  |                               256 = Error setting $iAlign
;                  |                               512 = Error setting $iLines
;                  |                               1024 = Error setting $aiDefaultSel
;                  |                               2048 = Error setting $sAddInfo
;                  |                               4096 = Error setting $sHelpText
;                  |                               8192 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 14 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The array called for $asList should be a single dimension array, with one List entry as a String, per array element.
;                  The array called for $aiDefaultSel should be a single dimension array, with one Integer value, corresponding to the position in the $asList array, per array element, to indicate which value(s) is/are default.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $asList, $sAddInfo.
; Related .......: _LOWriter_FormConListBoxData, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConListBoxGeneral(ByRef $oListBox, $sName = Null, $sLabel = Null, $iTxtDir = Null, $bEnabled = Null, $bReadOnly = Null, $iMouseScroll = Null, $iWidth = Null, $asList = Null, $iAlign = Null, $iLines = Null, $aiDefaultSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[14]

	If Not IsObj($oListBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oListBox) <> $LOW_FORM_CON_TYPE_LIST_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $bEnabled, $bReadOnly, $iMouseScroll, $iWidth, $asList, $iAlign, $iLines, $aiDefaultSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oListBox.Name(), $oListBox.Label(), $oListBox.WritingMode(), $oListBox.Enabled(), _
				$oListBox.ReadOnly(), $oListBox.MouseWheelBehavior(), Int($oListBox.Width() * 10), _ ; Multiply width by 10 to get Hundredths of a Millimeter value.
				$oListBox.StringItemList(), $oListBox.Align(), $oListBox.LineCount(), _
				$oListBox.DefaultSelection(), $oListBox.Tag(), $oListBox.HelpText(), $oListBox.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oListBox.Name = $sName
		$iError = ($oListBox.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oListBox.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oListBox.Label = $sLabel
		$iError = ($oListBox.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oListBox.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oListBox.WritingMode = $iTxtDir
		$iError = ($oListBox.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bEnabled = Default) Then
		$oListBox.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oListBox.Enabled = $bEnabled
		$iError = ($oListBox.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bReadOnly = Default) Then
		$oListBox.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oListBox.ReadOnly = $bReadOnly
		$iError = ($oListBox.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iMouseScroll = Default) Then
		$oListBox.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oListBox.MouseWheelBehavior = $iMouseScroll
		$iError = ($oListBox.MouseWheelBehavior = $iMouseScroll) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iWidth = Default) Then
		$oListBox.setPropertyToDefault("Width")

	ElseIf ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 100, 200000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oListBox.Width = Round($iWidth / 10) ; Divide Hundredths of a Millimeter value by 10 to obtain 10th MM.
		$iError = ($oListBox.Width() = Round($iWidth / 10)) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($asList = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default StringItemList.

	ElseIf ($asList <> Null) Then
		If Not IsArray($asList) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		For $i = 0 To UBound($asList) - 1
			If Not IsString($asList[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, $i)

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next
		$oListBox.StringItemList = $asList
		$iError = (UBound($oListBox.StringItemList()) = UBound($asList)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iAlign = Default) Then
		$oListBox.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oListBox.Align = $iAlign
		$iError = ($oListBox.Align() = $iAlign) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iLines = Default) Then
		$oListBox.setPropertyToDefault("LineCount")

	ElseIf ($iLines <> Null) Then
		If Not __LO_IntIsBetween($iLines, -2147483648, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oListBox.LineCount = $iLines
		$iError = ($oListBox.LineCount = $iLines) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($aiDefaultSel = Default) Then
		$iError = BitOR($iError, 2048) ; Can't Default Name.

	ElseIf ($aiDefaultSel <> Null) Then
		If Not IsArray($aiDefaultSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		For $i = 0 To UBound($aiDefaultSel) - 1
			If Not IsInt($aiDefaultSel[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, $i)
			If ($aiDefaultSel[$i] >= $oListBox.ItemCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, $i)

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next
		$oListBox.DefaultSelection = $aiDefaultSel
		$iError = (UBound($oListBox.DefaultSelection()) = UBound($aiDefaultSel)) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 4096) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oListBox.Tag = $sAddInfo
		$iError = ($oListBox.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($sHelpText = Default) Then
		$oListBox.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oListBox.HelpText = $sHelpText
		$iError = ($oListBox.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($sHelpURL = Default) Then
		$oListBox.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oListBox.HelpURL = $sHelpURL
		$iError = ($oListBox.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConListBoxGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConNumericFieldData
; Description ...: Set or Retrieve Table Control Numeric Field Data Properties.
; Syntax ........: _LOWriter_FormConTableConNumericFieldData(ByRef $oNumericField[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oNumericField       - [in/out] an object. A Table Control Numeric Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oNumericField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oNumericField not a Numeric Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConNumericFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConNumericFieldData(ByRef $oNumericField, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oNumericField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oNumericField) <> $LOW_FORM_CON_TYPE_NUMERIC_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oNumericField.DataField(), $oNumericField.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oNumericField.DataField = $sDataField
		$iError = ($oNumericField.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oNumericField.InputRequired = $bInputRequired
		$iError = ($oNumericField.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConNumericFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConNumericFieldGeneral
; Description ...: Set or Retrieve general Table Control Numeric Field properties.
; Syntax ........: _LOWriter_FormConTableConNumericFieldGeneral(ByRef $oNumericField[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $bStrict = Null[, $bEnabled = Null[, $bReadOnly = Null[, $iMouseScroll = Null[, $nMin = Null[, $nMax = Null[, $iIncr = Null[, $nDefault = Null[, $iDecimal = Null[, $bThousandsSep = Null[, $bSpin = Null[, $bRepeat = Null[, $iDelay = Null[, $iWidth = Null[, $iAlign = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oNumericField       - [in/out] an object. A Table Control Numeric Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bStrict             - [optional] a boolean value. Default is Null. If True, strict formatting is enabled.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $nMin                - [optional] a general number value. Default is Null. The minimum value the control can be set to.
;                  $nMax                - [optional] a general number value. Default is Null. The maximum value the control can be set to.
;                  $iIncr               - [optional] an integer value. Default is Null. The amount to Increase or Decrease the value by.
;                  $nDefault            - [optional] a general number value. Default is Null. The default value the control will be set to.
;                  $iDecimal            - [optional] an integer value (0-20). Default is Null. The amount of decimal accuracy.
;                  $bThousandsSep       - [optional] a boolean value. Default is Null. If True, a thousands separator will be added.
;                  $bSpin               - [optional] a boolean value. Default is Null. If True, the field will act as a spin button.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $iWidth              - [optional] an integer value (100-200000). Default is Null. The width of the Column tab, in Hundredths of a Millimeter (HMM).
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oNumericField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oNumericField not a Numeric Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bStrict not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $nMin not a Number.
;                  @Error 1 @Extended 11 Return 0 = $nMax not a Number.
;                  @Error 1 @Extended 12 Return 0 = $iIncr not an Integer.
;                  @Error 1 @Extended 13 Return 0 = $nDefault not a Number.
;                  @Error 1 @Extended 14 Return 0 = $iDecimal not an Integer, less than 0 or greater than 20.
;                  @Error 1 @Extended 15 Return 0 = $bThousandsSep not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $bSpin not a Boolean.
;                  @Error 1 @Extended 17 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 18 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 19 Return 0 = $iWidth not an Integer, less than 100 or greater than 200,000.
;                  @Error 1 @Extended 20 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 21 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 22 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 23 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 24 Return 0 = $sHelpURL not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.DateTime" Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.util.Time" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current Minimum Time.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve current Maximum Time.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify parent document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bStrict
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bReadOnly
;                  |                               64 = Error setting $iMouseScroll
;                  |                               128 = Error setting $nMin
;                  |                               256 = Error setting $nMax
;                  |                               512 = Error setting $iIncr
;                  |                               1024 = Error setting $nDefault
;                  |                               2048 = Error setting $iDecimal
;                  |                               4096 = Error setting $bThousandsSep
;                  |                               8192 = Error setting $bSpin
;                  |                               16384 = Error setting $bRepeat
;                  |                               32768 = Error setting $iDelay
;                  |                               65536 = Error setting $iWidth
;                  |                               131072 = Error setting $iAlign
;                  |                               262144 = Error setting $bHideSel
;                  |                               524288 = Error setting $sAddInfo
;                  |                               1048576 = Error setting $sHelpText
;                  |                               2097152 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 22 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $sAddInfo.
; Related .......: _LOWriter_FormConNumericFieldData, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConNumericFieldGeneral(ByRef $oNumericField, $sName = Null, $sLabel = Null, $iTxtDir = Null, $bStrict = Null, $bEnabled = Null, $bReadOnly = Null, $iMouseScroll = Null, $nMin = Null, $nMax = Null, $iIncr = Null, $nDefault = Null, $iDecimal = Null, $bThousandsSep = Null, $bSpin = Null, $bRepeat = Null, $iDelay = Null, $iWidth = Null, $iAlign = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[22]

	If Not IsObj($oNumericField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oNumericField) <> $LOW_FORM_CON_TYPE_NUMERIC_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $bStrict, $bEnabled, $bReadOnly, $iMouseScroll, $nMin, $nMax, $iIncr, $nDefault, $iDecimal, $bThousandsSep, $bSpin, $bRepeat, $iDelay, $iWidth, $iAlign, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oNumericField.Name(), $oNumericField.Label(), $oNumericField.WritingMode(), $oNumericField.StrictFormat(), _
				$oNumericField.Enabled(), $oNumericField.ReadOnly(), _
				$oNumericField.MouseWheelBehavior(), $oNumericField.ValueMin(), _
				$oNumericField.ValueMax(), $oNumericField.ValueStep(), $oNumericField.DefaultValue(), $oNumericField.DecimalAccuracy(), _
				$oNumericField.ShowThousandsSeparator(), $oNumericField.Spin(), $oNumericField.Repeat(), $oNumericField.RepeatDelay(), _
				Int($oNumericField.Width() * 10), _ ; Multiply width by 10 to get Hundredths of a Millimeter value.
				$oNumericField.Align(), $oNumericField.HideInactiveSelection(), $oNumericField.Tag(), _
				$oNumericField.HelpText(), $oNumericField.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oNumericField.Name = $sName
		$iError = ($oNumericField.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oNumericField.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oNumericField.Label = $sLabel
		$iError = ($oNumericField.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oNumericField.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oNumericField.WritingMode = $iTxtDir
		$iError = ($oNumericField.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bStrict = Default) Then
		$oNumericField.setPropertyToDefault("StrictFormat")

	ElseIf ($bStrict <> Null) Then
		If Not IsBool($bStrict) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oNumericField.StrictFormat = $bStrict
		$iError = ($oNumericField.StrictFormat() = $bStrict) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oNumericField.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oNumericField.Enabled = $bEnabled
		$iError = ($oNumericField.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bReadOnly = Default) Then
		$oNumericField.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oNumericField.ReadOnly = $bReadOnly
		$iError = ($oNumericField.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iMouseScroll = Default) Then
		$oNumericField.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oNumericField.MouseWheelBehavior = $iMouseScroll
		$iError = ($oNumericField.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($nMin = Default) Then
		$oNumericField.setPropertyToDefault("ValueMin")

	ElseIf ($nMin <> Null) Then
		If Not IsNumber($nMin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oNumericField.ValueMin = $nMin
		$iError = ($oNumericField.ValueMin() = $nMin) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($nMax = Default) Then
		$oNumericField.setPropertyToDefault("ValueMax")

	ElseIf ($nMax <> Null) Then
		If Not IsNumber($nMax) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oNumericField.ValueMax = $nMax
		$iError = ($oNumericField.ValueMax() = $nMax) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iIncr = Default) Then
		$oNumericField.setPropertyToDefault("ValueStep")

	ElseIf ($iIncr <> Null) Then
		If Not IsInt($iIncr) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oNumericField.ValueStep = $iIncr
		$iError = ($oNumericField.ValueStep() = $iIncr) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($nDefault = Default) Then
		$oNumericField.setPropertyToDefault("DefaultValue")

	ElseIf ($nDefault <> Null) Then
		If Not IsNumber($nDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oNumericField.DefaultValue = $nDefault
		$iError = ($oNumericField.DefaultValue() = $nDefault) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($iDecimal = Default) Then
		$oNumericField.setPropertyToDefault("DecimalAccuracy")

	ElseIf ($iDecimal <> Null) Then
		If Not __LO_IntIsBetween($iDecimal, 0, 20) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oNumericField.DecimalAccuracy = $iDecimal
		$iError = ($oNumericField.DecimalAccuracy() = $iDecimal) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($bThousandsSep = Default) Then
		$oNumericField.setPropertyToDefault("ShowThousandsSeparator")

	ElseIf ($bThousandsSep <> Null) Then
		If Not IsBool($bThousandsSep) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oNumericField.ShowThousandsSeparator = $bThousandsSep
		$iError = ($oNumericField.ShowThousandsSeparator() = $bThousandsSep) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($bSpin = Default) Then
		$oNumericField.setPropertyToDefault("Spin")

	ElseIf ($bSpin <> Null) Then
		If Not IsBool($bSpin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oNumericField.Spin = $bSpin
		$iError = ($oNumericField.Spin() = $bSpin) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($bRepeat = Default) Then
		$oNumericField.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oNumericField.Repeat = $bRepeat
		$iError = ($oNumericField.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iDelay = Default) Then
		$oNumericField.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oNumericField.RepeatDelay = $iDelay
		$iError = ($oNumericField.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($iWidth = Default) Then
		$oNumericField.setPropertyToDefault("Width")

	ElseIf ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 100, 200000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oNumericField.Width = Round($iWidth / 10) ; Divide Hundredths of a Millimeter value by 10 to obtain 10th MM.
		$iError = ($oNumericField.Width() = Round($iWidth / 10)) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($iAlign = Default) Then
		$oNumericField.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oNumericField.Align = $iAlign
		$iError = ($oNumericField.Align() = $iAlign) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($bHideSel = Default) Then
		$oNumericField.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oNumericField.HideInactiveSelection = $bHideSel
		$iError = ($oNumericField.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 524288) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oNumericField.Tag = $sAddInfo
		$iError = ($oNumericField.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($sHelpText = Default) Then
		$oNumericField.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oNumericField.HelpText = $sHelpText
		$iError = ($oNumericField.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($sHelpURL = Default) Then
		$oNumericField.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oNumericField.HelpURL = $sHelpURL
		$iError = ($oNumericField.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConNumericFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConPatternFieldData
; Description ...: Set or Retrieve Table Control Pattern Field Data Properties.
; Syntax ........: _LOWriter_FormConTableConPatternFieldData(ByRef $oPatternField[, $sDataField = Null[, $bEmptyIsNull = Null[, $bInputRequired = Null[, $bFilter = Null]]]])
; Parameters ....: $oPatternField       - [in/out] an object. A Table Control Pattern Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bEmptyIsNull        - [optional] a boolean value. Default is Null. If True, an empty string will be treated as a Null value.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
;                  $bFilter             - [optional] a boolean value. Default is Null. If True, filter proposal is active.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPatternField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oPatternField not a Pattern Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bEmptyIsNull not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bInputRequired not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bFilter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bEmptyIsNull
;                  |                               4 = Error setting $bInputRequired
;                  |                               8 = Error setting $bFilter
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConPatternFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConPatternFieldData(ByRef $oPatternField, $sDataField = Null, $bEmptyIsNull = Null, $bInputRequired = Null, $bFilter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[4]

	If Not IsObj($oPatternField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oPatternField) <> $LOW_FORM_CON_TYPE_PATTERN_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bEmptyIsNull, $bInputRequired, $bFilter) Then
		__LO_ArrayFill($avControl, $oPatternField.DataField(), $oPatternField.ConvertEmptyToNull(), $oPatternField.InputRequired(), _
				$oPatternField.UseFilterValueProposal())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPatternField.DataField = $sDataField
		$iError = ($oPatternField.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bEmptyIsNull <> Null) Then
		If Not IsBool($bEmptyIsNull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPatternField.ConvertEmptyToNull = $bEmptyIsNull
		$iError = ($oPatternField.ConvertEmptyToNull() = $bEmptyIsNull) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPatternField.InputRequired = $bInputRequired
		$iError = ($oPatternField.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bFilter <> Null) Then
		If Not IsBool($bFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPatternField.UseFilterValueProposal = $bFilter
		$iError = ($oPatternField.UseFilterValueProposal() = $bFilter) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConPatternFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConPatternFieldGeneral
; Description ...: Set or Retrieve general Table Control Pattern Field properties.
; Syntax ........: _LOWriter_FormConTableConPatternFieldGeneral(ByRef $oPatternField[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $iMaxLen = Null[, $sEditMask = Null[, $sLiteralMask = Null[, $bStrict = Null[, $bEnabled = Null[, $bReadOnly = Null[, $iMouseScroll = Null[, $iWidth = Null[, $sDefaultTxt = Null[, $iAlign = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]])
; Parameters ....: $oPatternField       - [in/out] an object. A Table Control Pattern Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iMaxLen             - [optional] an integer value (-1-2147483647). Default is Null. The maximum text length that the Pattern field will accept.
;                  $sEditMask           - [optional] a string value. Default is Null. The edit mask of the field.
;                  $sLiteralMask        - [optional] a string value. Default is Null. The literal mask of the field.
;                  $bStrict             - [optional] a boolean value. Default is Null. If True, strict formatting is enabled.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iWidth              - [optional] an integer value (100-200000). Default is Null. The width of the Column tab, in Hundredths of a Millimeter (HMM).
;                  $sDefaultTxt         - [optional] a string value. Default is Null. The default text to display in the field.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPatternField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oPatternField not a Pattern Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iMaxLen not an Integer, less than -1 or greater than 2147483647.
;                  @Error 1 @Extended 7 Return 0 = $sEditMask
;                  @Error 1 @Extended 8 Return 0 = $sLiteralMask not a String.
;                  @Error 1 @Extended 9 Return 0 = $bStrict not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $iWidth not an Integer, less than 100 or greater than 200,000.
;                  @Error 1 @Extended 14 Return 0 = $sDefaultTxt not a String.
;                  @Error 1 @Extended 15 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 16 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 17 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 18 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 19 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $iMaxLen
;                  |                               16 = Error setting $sEditMask
;                  |                               32 = Error setting $sLiteralMask
;                  |                               64 = Error setting $bStrict
;                  |                               128 = Error setting $bEnabled
;                  |                               256 = Error setting $bReadOnly
;                  |                               512 = Error setting $iMouseScroll
;                  |                               1024 = Error setting $iWidth
;                  |                               2048 = Error setting $sDefaultTxt
;                  |                               4096 = Error setting $iAlign
;                  |                               8192 = Error setting $bHideSel
;                  |                               16384 = Error setting $sAddInfo
;                  |                               32768 = Error setting $sHelpText
;                  |                               65536 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 17 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $sDefaultTxt, $sAddInfo.
; Related .......: _LOWriter_FormConPatternFieldData, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConPatternFieldGeneral(ByRef $oPatternField, $sName = Null, $sLabel = Null, $iTxtDir = Null, $iMaxLen = Null, $sEditMask = Null, $sLiteralMask = Null, $bStrict = Null, $bEnabled = Null, $bReadOnly = Null, $iMouseScroll = Null, $iWidth = Null, $sDefaultTxt = Null, $iAlign = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[17]

	If Not IsObj($oPatternField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oPatternField) <> $LOW_FORM_CON_TYPE_PATTERN_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $iMaxLen, $sEditMask, $sLiteralMask, $bStrict, $bEnabled, $bReadOnly, $iMouseScroll, $iWidth, $sDefaultTxt, $iAlign, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oPatternField.Name(), $oPatternField.Label(), $oPatternField.WritingMode(), _
				$oPatternField.MaxTextLen(), $oPatternField.EditMask(), $oPatternField.LiteralMask(), $oPatternField.StrictFormat(), _
				$oPatternField.Enabled(), $oPatternField.ReadOnly(), _
				$oPatternField.MouseWheelBehavior(), Int($oPatternField.Width() * 10), _ ; Multiply width by 10 to get Hundredths of a Millimeter value.
				$oPatternField.DefaultText(), _
				$oPatternField.Align(), $oPatternField.HideInactiveSelection(), $oPatternField.Tag(), _
				$oPatternField.HelpText(), $oPatternField.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oPatternField.Name = $sName
		$iError = ($oPatternField.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oPatternField.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oPatternField.Label = $sLabel
		$iError = ($oPatternField.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oPatternField.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oPatternField.WritingMode = $iTxtDir
		$iError = ($oPatternField.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iMaxLen = Default) Then
		$oPatternField.setPropertyToDefault("MaxTextLen")

	ElseIf ($iMaxLen <> Null) Then
		If Not __LO_IntIsBetween($iMaxLen, -1, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oPatternField.MaxTextLen = $iMaxLen
		$iError = ($oPatternField.MaxTextLen = $iMaxLen) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sEditMask = Default) Then
		$oPatternField.setPropertyToDefault("EditMask")

	ElseIf ($sEditMask <> Null) Then
		If Not IsString($sEditMask) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oPatternField.EditMask = $sEditMask
		$iError = ($oPatternField.EditMask() = $sEditMask) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($sLiteralMask = Default) Then
		$oPatternField.setPropertyToDefault("LiteralMask")

	ElseIf ($sLiteralMask <> Null) Then
		If Not IsString($sLiteralMask) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oPatternField.LiteralMask = $sLiteralMask
		$iError = ($oPatternField.LiteralMask() = $sLiteralMask) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bStrict = Default) Then
		$oPatternField.setPropertyToDefault("StrictFormat")

	ElseIf ($bStrict <> Null) Then
		If Not IsBool($bStrict) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oPatternField.StrictFormat = $bStrict
		$iError = ($oPatternField.StrictFormat() = $bStrict) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bEnabled = Default) Then
		$oPatternField.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oPatternField.Enabled = $bEnabled
		$iError = ($oPatternField.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bReadOnly = Default) Then
		$oPatternField.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oPatternField.ReadOnly = $bReadOnly
		$iError = ($oPatternField.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iMouseScroll = Default) Then
		$oPatternField.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oPatternField.MouseWheelBehavior = $iMouseScroll
		$iError = ($oPatternField.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iWidth = Default) Then
		$oPatternField.setPropertyToDefault("Width")

	ElseIf ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 100, 200000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oPatternField.Width = Round($iWidth / 10) ; Divide Hundredths of a Millimeter value by 10 to obtain 10th MM.
		$iError = ($oPatternField.Width() = Round($iWidth / 10)) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($sDefaultTxt = Default) Then
		$iError = BitOR($iError, 16384) ; Can't Default DefaultText.

	ElseIf ($sDefaultTxt <> Null) Then
		If Not IsString($sDefaultTxt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oPatternField.DefaultText = $sDefaultTxt
		$iError = ($oPatternField.DefaultText() = $sDefaultTxt) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($iAlign = Default) Then
		$oPatternField.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oPatternField.Align = $iAlign
		$iError = ($oPatternField.Align() = $iAlign) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($bHideSel = Default) Then
		$oPatternField.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oPatternField.HideInactiveSelection = $bHideSel
		$iError = ($oPatternField.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 16384) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oPatternField.Tag = $sAddInfo
		$iError = ($oPatternField.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($sHelpText = Default) Then
		$oPatternField.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oPatternField.HelpText = $sHelpText
		$iError = ($oPatternField.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($sHelpURL = Default) Then
		$oPatternField.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oPatternField.HelpURL = $sHelpURL
		$iError = ($oPatternField.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConPatternFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConTextBoxData
; Description ...: Set or Retrieve Table Control Text Box Data Properties.
; Syntax ........: _LOWriter_FormConTableConTextBoxData(ByRef $oTextBox[, $sDataField = Null[, $bEmptyIsNull = Null[, $bInputRequired = Null[, $bFilter = Null]]]])
; Parameters ....: $oTextBox            - [in/out] an object. A Table Control Text Box Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bEmptyIsNull        - [optional] a boolean value. Default is Null. If True, an empty string will be treated as a Null value.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
;                  $bFilter             - [optional] a boolean value. Default is Null. If True, filter proposal is active.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTextBox not a Text Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bEmptyIsNull not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bInputRequired not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bFilter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bEmptyIsNull
;                  |                               4 = Error setting $bInputRequired
;                  |                               8 = Error setting $bFilter
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConTextBoxGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConTextBoxData(ByRef $oTextBox, $sDataField = Null, $bEmptyIsNull = Null, $bInputRequired = Null, $bFilter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[4]

	If Not IsObj($oTextBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTextBox) <> $LOW_FORM_CON_TYPE_TEXT_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bEmptyIsNull, $bInputRequired, $bFilter) Then
		__LO_ArrayFill($avControl, $oTextBox.DataField(), $oTextBox.ConvertEmptyToNull(), $oTextBox.InputRequired(), _
				$oTextBox.UseFilterValueProposal())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTextBox.DataField = $sDataField
		$iError = ($oTextBox.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bEmptyIsNull <> Null) Then
		If Not IsBool($bEmptyIsNull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTextBox.ConvertEmptyToNull = $bEmptyIsNull
		$iError = ($oTextBox.ConvertEmptyToNull() = $bEmptyIsNull) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTextBox.InputRequired = $bInputRequired
		$iError = ($oTextBox.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bFilter <> Null) Then
		If Not IsBool($bFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTextBox.UseFilterValueProposal = $bFilter
		$iError = ($oTextBox.UseFilterValueProposal() = $bFilter) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConTextBoxData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConTextBoxGeneral
; Description ...: Set or Retrieve general Table Control Textbox control properties.
; Syntax ........: _LOWriter_FormConTableConTextBoxGeneral(ByRef $oTextBox[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $iMaxLen = Null[, $bEnabled = Null[, $bReadOnly = Null[, $iWidth = Null[, $sDefaultTxt = Null[, $iAlign = Null[, $bMultiLine = Null[, $bEndWithLF = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]])
; Parameters ....: $oTextBox            - [in/out] an object. A Table Control Textbox Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control Name.
;                  $sLabel              - [optional] a string value. Default is Null. The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iMaxLen             - [optional] an integer value (-1-2147483647). Default is Null. The max length of text that can be entered. 0 = unlimited.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $iWidth              - [optional] an integer value (100-200000). Default is Null. The width of the Column tab, in Hundredths of a Millimeter (HMM).
;                  $sDefaultTxt         - [optional] a string value. Default is Null. The default text to display in the control.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bMultiLine          - [optional] a boolean value. Default is Null. If True, the text may contain multiple lines.
;                  $bEndWithLF          - [optional] a boolean value. Default is Null. If True, the line ends will be a Line-Feed type, else a Carriage-Return plus a Line-Feed.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTextBox not a Text Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iMaxLen not an Integer, less than -1 or greater than 2147483647.
;                  @Error 1 @Extended 7 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iWidth not an Integer, less than 100 or greater than 200,000.
;                  @Error 1 @Extended 10 Return 0 = $sDefaultTxt not a String.
;                  @Error 1 @Extended 11 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 12 Return 0 = $bMultiLine not a Boolean.
;                  @Error 1 @Extended 13 Return 0 = $bEndWithLF not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 16 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 17 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $iMaxLen
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bReadOnly
;                  |                               64 = Error setting $iWidth
;                  |                               128 = Error setting $sDefaultTxt
;                  |                               256 = Error setting $iAlign
;                  |                               512 = Error setting $bMultiLine
;                  |                               1024 = Error setting $bEndWithLF
;                  |                               2048 = Error setting $bHideSel
;                  |                               4096 = Error setting $sAddInfo
;                  |                               8192 = Error setting $sHelpText
;                  |                               16384 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 15 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $sDefaultTxt, $sAddInfo.
; Related .......: _LOWriter_FormConTextBoxData, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConTextBoxGeneral(ByRef $oTextBox, $sName = Null, $sLabel = Null, $iTxtDir = Null, $iMaxLen = Null, $bEnabled = Null, $bReadOnly = Null, $iWidth = Null, $sDefaultTxt = Null, $iAlign = Null, $bMultiLine = Null, $bEndWithLF = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[15]
	Local Const $__LOW_FORM_CONTROL_LINE_END_CR = 0, $__LOW_FORM_CONTROL_LINE_END_LF = 1, $__LOW_FORM_CONTROL_LINE_END_CRLF = 2 ; "com.sun.star.awt.LineEndFormat"
	#forceref $__LOW_FORM_CONTROL_LINE_END_CR

	If Not IsObj($oTextBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTextBox) <> $LOW_FORM_CON_TYPE_TEXT_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $iMaxLen, $bEnabled, $bReadOnly, $iWidth, $sDefaultTxt, $iAlign, $bMultiLine, $bEndWithLF, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oTextBox.Name(), $oTextBox.Label(), $oTextBox.WritingMode(), $oTextBox.MaxTextLen(), _
				$oTextBox.Enabled(), $oTextBox.ReadOnly(), _
				Int($oTextBox.Width() * 10), _ ; Multiply width by 10 to get Hundredths of a Millimeter value.
				$oTextBox.DefaultText(), $oTextBox.Align(), $oTextBox.MultiLine(), _
				(($oTextBox.LineEndFormat() = $__LOW_FORM_CONTROL_LINE_END_LF) ? (True) : (False)), _ ; Line Ending format
				$oTextBox.HideInactiveSelection(), $oTextBox.Tag(), $oTextBox.HelpText(), $oTextBox.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTextBox.Name = $sName
		$iError = ($oTextBox.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oTextBox.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTextBox.Label = $sLabel
		$iError = ($oTextBox.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oTextBox.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTextBox.WritingMode = $iTxtDir
		$iError = ($oTextBox.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iMaxLen = Default) Then
		$oTextBox.setPropertyToDefault("MaxTextLen")

	ElseIf ($iMaxLen <> Null) Then
		If Not __LO_IntIsBetween($iMaxLen, -1, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTextBox.MaxTextLen = $iMaxLen
		$iError = ($oTextBox.MaxTextLen() = $iMaxLen) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oTextBox.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oTextBox.Enabled = $bEnabled
		$iError = ($oTextBox.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bReadOnly = Default) Then
		$oTextBox.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oTextBox.ReadOnly = $bReadOnly
		$iError = ($oTextBox.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iWidth = Default) Then
		$oTextBox.setPropertyToDefault("Width")

	ElseIf ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 100, 200000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oTextBox.Width = Round($iWidth / 10) ; Divide Hundredths of a Millimeter value by 10 to obtain 10th MM.
		$iError = ($oTextBox.Width() = Round($iWidth / 10)) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($sDefaultTxt = Default) Then
		$iError = BitOR($iError, 128) ; Can't Default DefaultText.

	ElseIf ($sDefaultTxt <> Null) Then
		If Not IsString($sDefaultTxt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oTextBox.DefaultText = $sDefaultTxt
		$iError = ($oTextBox.DefaultText() = $sDefaultTxt) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iAlign = Default) Then
		$oTextBox.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oTextBox.Align = $iAlign
		$iError = ($oTextBox.Align() = $iAlign) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bMultiLine = Default) Then
		$oTextBox.setPropertyToDefault("MultiLine")

	ElseIf ($bMultiLine <> Null) Then
		If Not IsBool($bMultiLine) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oTextBox.MultiLine = $bMultiLine
		$iError = ($oTextBox.MultiLine = $bMultiLine) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($bEndWithLF = Default) Then
		$oTextBox.setPropertyToDefault("LineEndFormat")

	ElseIf ($bEndWithLF <> Null) Then
		If Not IsBool($bEndWithLF) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		If $bEndWithLF Then
			$oTextBox.LineEndFormat = $__LOW_FORM_CONTROL_LINE_END_LF
			$iError = ($oTextBox.LineEndFormat = $__LOW_FORM_CONTROL_LINE_END_LF) ? ($iError) : (BitOR($iError, 1024))

		Else
			$oTextBox.LineEndFormat = $__LOW_FORM_CONTROL_LINE_END_CRLF
			$iError = ($oTextBox.LineEndFormat = $__LOW_FORM_CONTROL_LINE_END_CRLF) ? ($iError) : (BitOR($iError, 1024))
		EndIf
	EndIf

	If ($bHideSel = Default) Then
		$oTextBox.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oTextBox.HideInactiveSelection = $bHideSel
		$iError = ($oTextBox.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 4096) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oTextBox.Tag = $sAddInfo
		$iError = ($oTextBox.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($sHelpText = Default) Then
		$oTextBox.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oTextBox.HelpText = $sHelpText
		$iError = ($oTextBox.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($sHelpURL = Default) Then
		$oTextBox.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oTextBox.HelpURL = $sHelpURL
		$iError = ($oTextBox.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConTextBoxGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConTimeFieldData
; Description ...: Set or Retrieve Table Control Time Field Data Properties.
; Syntax ........: _LOWriter_FormConTableConTimeFieldData(ByRef $oTimeField[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oTimeField          - [in/out] an object. A Table Control Time Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTimeField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTimeField not a Time Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConTimeFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConTimeFieldData(ByRef $oTimeField, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oTimeField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTimeField) <> $LOW_FORM_CON_TYPE_TIME_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oTimeField.DataField(), $oTimeField.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTimeField.DataField = $sDataField
		$iError = ($oTimeField.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTimeField.InputRequired = $bInputRequired
		$iError = ($oTimeField.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConTimeFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTableConTimeFieldGeneral
; Description ...: Set or Retrieve general Table Control Time Field properties.
; Syntax ........: _LOWriter_FormConTableConTimeFieldGeneral(ByRef $oTimeField[, $sName = Null[, $sLabel = Null[, $iTxtDir = Null[, $bStrict = Null[, $bEnabled = Null[, $bReadOnly = Null[, $iMouseScroll = Null[, $tTimeMin = Null[, $tTimeMax = Null[, $iFormat = Null[, $tTimeDefault = Null[, $bSpin = Null[, $bRepeat = Null[, $iDelay = Null[, $iWidth = Null[, $iAlign = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oTimeField          - [in/out] an object. A Table Control Time Field Control object returned by a previous _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConTableConColumnsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sLabel              - [optional] a string value. Default is Null.The control's label to display.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bStrict             - [optional] a boolean value. Default is Null. If True, strict formatting is enabled.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $tTimeMin            - [optional] a dll struct value. Default is Null. The minimum time	 allowed to be entered, created previously by _LOWriter_DateStructCreate.
;                  $tTimeMax            - [optional] a dll struct value. Default is Null. The maximum time	 allowed to be entered, created previously by _LOWriter_DateStructCreate.
;                  $iFormat             - [optional] an integer value (0-5). Default is Null. The Time Format to display the content in. See Constants $LOW_FORM_CON_TIME_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $tTimeDefault        - [optional] a dll struct value. Default is Null. The Default time to display, created previously by _LOWriter_DateStructCreate.
;                  $bSpin               - [optional] a boolean value. Default is Null. If True, the field will act as a spin button.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $iWidth              - [optional] an integer value (100-200000). Default is Null. The width of the Column tab, in Hundredths of a Millimeter (HMM).
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTimeField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTimeField not a Time Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bStrict not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $tTimeMin not an Object.
;                  @Error 1 @Extended 11 Return 0 = $tTimeMax not an Object.
;                  @Error 1 @Extended 12 Return 0 = $iFormat not an Integer, less than 0 or greater than 5. See Constants $LOW_FORM_CON_TIME_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $tTimeDefault not an Object.
;                  @Error 1 @Extended 14 Return 0 = $bSpin not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 17 Return 0 = $iWidth not an Integer, less than 100 or greater than 200,000.
;                  @Error 1 @Extended 18 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 19 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 20 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 21 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 22 Return 0 = $sHelpURL not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.DateTime" Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.util.Time" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current Minimum Time.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve current Maximum Time.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bStrict
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bReadOnly
;                  |                               64 = Error setting $iMouseScroll
;                  |                               128 = Error setting $tTimeMin
;                  |                               256 = Error setting $tTimeMax
;                  |                               512 = Error setting $iFormat
;                  |                               1024 = Error setting $tTimeDefault
;                  |                               2048 = Error setting $bSpin
;                  |                               4096 = Error setting $bRepeat
;                  |                               8192 = Error setting $iDelay
;                  |                               16384 = Error setting $iWidth
;                  |                               32768 = Error setting $iAlign
;                  |                               65536 = Error setting $bHideSel
;                  |                               131072 = Error setting $sAddInfo
;                  |                               262144 = Error setting $sHelpText
;                  |                               524288 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 20 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $mFont, $sAddInfo.
; Related .......: _LOWriter_FormatKeyCreate, _LOWriter_FormatKeysGetList, _LOWriter_FormConTimeFieldData, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTableConTimeFieldGeneral(ByRef $oTimeField, $sName = Null, $sLabel = Null, $iTxtDir = Null, $bStrict = Null, $bEnabled = Null, $bReadOnly = Null, $iMouseScroll = Null, $tTimeMin = Null, $tTimeMax = Null, $iFormat = Null, $tTimeDefault = Null, $bSpin = Null, $bRepeat = Null, $iDelay = Null, $iWidth = Null, $iAlign = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tTime, $tCurMin, $tCurMax, $tCurDefault
	Local $avControl[20]

	If Not IsObj($oTimeField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTimeField) <> $LOW_FORM_CON_TYPE_TIME_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $iTxtDir, $bStrict, $bEnabled, $bReadOnly, $iMouseScroll, $tTimeMin, $tTimeMax, $iFormat, $tTimeDefault, $bSpin, $bRepeat, $iDelay, $iWidth, $iAlign, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		$tTime = $oTimeField.TimeMin()
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$tCurMin = __LO_CreateStruct("com.sun.star.util.DateTime")
		If Not IsObj($tCurMin) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tCurMin.Hours = $tTime.Hours()
		$tCurMin.Minutes = $tTime.Minutes()
		$tCurMin.Seconds = $tTime.Seconds()
		$tCurMin.NanoSeconds = $tTime.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tCurMin.IsUTC = $tTime.IsUTC()

		$tTime = $oTimeField.TimeMax()
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$tCurMax = __LO_CreateStruct("com.sun.star.util.DateTime")
		If Not IsObj($tCurMax) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tCurMax.Hours = $tTime.Hours()
		$tCurMax.Minutes = $tTime.Minutes()
		$tCurMax.Seconds = $tTime.Seconds()
		$tCurMax.NanoSeconds = $tTime.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tCurMax.IsUTC = $tTime.IsUTC()

		$tTime = $oTimeField.DefaultTime() ; Default time is Null when not set.
		If IsObj($tTime) Then
			$tCurDefault = __LO_CreateStruct("com.sun.star.util.DateTime")
			If Not IsObj($tCurDefault) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tCurDefault.Hours = $tTime.Hours()
			$tCurDefault.Minutes = $tTime.Minutes()
			$tCurDefault.Seconds = $tTime.Seconds()
			$tCurDefault.NanoSeconds = $tTime.NanoSeconds()
			If __LO_VersionCheck(4.1) Then $tCurDefault.IsUTC = $tTime.IsUTC()

		Else
			$tCurDefault = $tTime
		EndIf

		__LO_ArrayFill($avControl, $oTimeField.Name(), $oTimeField.Label(), $oTimeField.WritingMode(), $oTimeField.StrictFormat(), _
				$oTimeField.Enabled(), $oTimeField.ReadOnly(), $oTimeField.MouseWheelBehavior(), _
				$tCurMin, $tCurMax, $oTimeField.TimeFormat(), $tCurDefault, $oTimeField.Spin(), _
				$oTimeField.Repeat(), $oTimeField.RepeatDelay(), Int($oTimeField.Width() * 10), _ ; Multiply width by 10 to get Hundredths of a Millimeter value.
				$oTimeField.Align(), $oTimeField.HideInactiveSelection(), $oTimeField.Tag(), $oTimeField.HelpText(), $oTimeField.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTimeField.Name = $sName
		$iError = ($oTimeField.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel = Default) Then
		$oTimeField.setPropertyToDefault("Label")

	ElseIf ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTimeField.Label = $sLabel
		$iError = ($oTimeField.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oTimeField.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTimeField.WritingMode = $iTxtDir
		$iError = ($oTimeField.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bStrict = Default) Then
		$oTimeField.setPropertyToDefault("StrictFormat")

	ElseIf ($bStrict <> Null) Then
		If Not IsBool($bStrict) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTimeField.StrictFormat = $bStrict
		$iError = ($oTimeField.StrictFormat() = $bStrict) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oTimeField.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oTimeField.Enabled = $bEnabled
		$iError = ($oTimeField.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bReadOnly = Default) Then
		$oTimeField.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oTimeField.ReadOnly = $bReadOnly
		$iError = ($oTimeField.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iMouseScroll = Default) Then
		$oTimeField.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oTimeField.MouseWheelBehavior = $iMouseScroll
		$iError = ($oTimeField.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($tTimeMin = Default) Then
		$oTimeField.setPropertyToDefault("TimeMin")

	ElseIf ($tTimeMin <> Null) Then
		If Not IsObj($tTimeMin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tTime = __LO_CreateStruct("com.sun.star.util.Time")
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tTime.Hours = $tTimeMin.Hours()
		$tTime.Minutes = $tTimeMin.Minutes()
		$tTime.Seconds = $tTimeMin.Seconds()
		$tTime.NanoSeconds = $tTimeMin.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tTime.IsUTC = $tTimeMin.IsUTC()

		$oTimeField.TimeMin = $tTime
		$iError = (__LOWriter_DateStructCompare($oTimeField.TimeMin(), $tTime, False, True)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($tTimeMax = Default) Then
		$oTimeField.setPropertyToDefault("TimeMax")

	ElseIf ($tTimeMax <> Null) Then
		If Not IsObj($tTimeMax) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tTime = __LO_CreateStruct("com.sun.star.util.Time")
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tTime.Hours = $tTimeMax.Hours()
		$tTime.Minutes = $tTimeMax.Minutes()
		$tTime.Seconds = $tTimeMax.Seconds()
		$tTime.NanoSeconds = $tTimeMax.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tTime.IsUTC = $tTimeMax.IsUTC()

		$oTimeField.TimeMax = $tTime
		$iError = (__LOWriter_DateStructCompare($oTimeField.TimeMax(), $tTime, False, True)) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iFormat = Default) Then
		$oTimeField.setPropertyToDefault("TimeFormat")

	ElseIf ($iFormat <> Null) Then
		If Not __LO_IntIsBetween($iFormat, $LOW_FORM_CON_TIME_FRMT_24_SHORT, $LOW_FORM_CON_TIME_FRMT_DURATION_LONG) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oTimeField.TimeFormat = $iFormat
		$iError = ($oTimeField.TimeFormat() = $iFormat) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($tTimeDefault = Default) Then
		$oTimeField.setPropertyToDefault("DefaultTime")

	ElseIf ($tTimeDefault <> Null) Then
		If Not IsObj($tTimeDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$tTime = __LO_CreateStruct("com.sun.star.util.Time")
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tTime.Hours = $tTimeDefault.Hours()
		$tTime.Minutes = $tTimeDefault.Minutes()
		$tTime.Seconds = $tTimeDefault.Seconds()
		$tTime.NanoSeconds = $tTimeDefault.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tTime.IsUTC = $tTimeDefault.IsUTC()

		$oTimeField.DefaultTime = $tTime
		$iError = (__LOWriter_DateStructCompare($oTimeField.DefaultTime(), $tTime, False, True)) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($bSpin = Default) Then
		$oTimeField.setPropertyToDefault("Spin")

	ElseIf ($bSpin <> Null) Then
		If Not IsBool($bSpin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oTimeField.Spin = $bSpin
		$iError = ($oTimeField.Spin() = $bSpin) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($bRepeat = Default) Then
		$oTimeField.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oTimeField.Repeat = $bRepeat
		$iError = ($oTimeField.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iDelay = Default) Then
		$oTimeField.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oTimeField.RepeatDelay = $iDelay
		$iError = ($oTimeField.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iWidth = Default) Then
		$oTimeField.setPropertyToDefault("Width")

	ElseIf ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 100, 200000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oTimeField.Width = Round($iWidth / 10) ; Divide Hundredths of a Millimeter value by 10 to obtain 10th MM.
		$iError = ($oTimeField.Width() = Round($iWidth / 10)) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iAlign = Default) Then
		$oTimeField.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oTimeField.Align = $iAlign
		$iError = ($oTimeField.Align() = $iAlign) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bHideSel = Default) Then
		$oTimeField.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oTimeField.HideInactiveSelection = $bHideSel
		$iError = ($oTimeField.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 131072) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oTimeField.Tag = $sAddInfo
		$iError = ($oTimeField.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($sHelpText = Default) Then
		$oTimeField.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oTimeField.HelpText = $sHelpText
		$iError = ($oTimeField.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($sHelpURL = Default) Then
		$oTimeField.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		$oTimeField.HelpURL = $sHelpURL
		$iError = ($oTimeField.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTableConTimeFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTextBoxCreateTextCursor
; Description ...: Create a Text Cursor in a Text Box to add text etc.
; Syntax ........: _LOWriter_FormConTextBoxCreateTextCursor(ByRef $oTextBox)
; Parameters ....: $oTextBox            - [in/out] an object. A Text Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTextBox not a Text Box Control.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Text Cursor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify control.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the Text Cursor object created in the Text Box.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: I am unable to format text in a Text Box (even manually), even though it is supposed to be possible. Thus formatting may or may not work.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTextBoxCreateTextCursor(ByRef $oTextBox)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCursor

	If Not IsObj($oTextBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTextBox) <> $LOW_FORM_CON_TYPE_TEXT_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oCursor = $oTextBox.Control.createTextCursor()
	If Not IsObj($oTextBox) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCursor)
EndFunc   ;==>_LOWriter_FormConTextBoxCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTextBoxData
; Description ...: Set or Retrieve Text Box Data Properties.
; Syntax ........: _LOWriter_FormConTextBoxData(ByRef $oTextBox[, $sDataField = Null[, $bEmptyIsNull = Null[, $bInputRequired = Null[, $bFilter = Null]]]])
; Parameters ....: $oTextBox            - [in/out] an object. A Text Box Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bEmptyIsNull        - [optional] a boolean value. Default is Null. If True, an empty string will be treated as a Null value.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
;                  $bFilter             - [optional] a boolean value. Default is Null. If True, filter proposal is active.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTextBox not a Text Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bEmptyIsNull not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bInputRequired not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bFilter not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bEmptyIsNull
;                  |                               4 = Error setting $bInputRequired
;                  |                               8 = Error setting $bFilter
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConTextBoxGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTextBoxData(ByRef $oTextBox, $sDataField = Null, $bEmptyIsNull = Null, $bInputRequired = Null, $bFilter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[4]

	If Not IsObj($oTextBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTextBox) <> $LOW_FORM_CON_TYPE_TEXT_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bEmptyIsNull, $bInputRequired, $bFilter) Then
		__LO_ArrayFill($avControl, $oTextBox.Control.DataField(), $oTextBox.Control.ConvertEmptyToNull(), $oTextBox.Control.InputRequired(), _
				$oTextBox.Control.UseFilterValueProposal())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTextBox.Control.DataField = $sDataField
		$iError = ($oTextBox.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bEmptyIsNull <> Null) Then
		If Not IsBool($bEmptyIsNull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTextBox.Control.ConvertEmptyToNull = $bEmptyIsNull
		$iError = ($oTextBox.Control.ConvertEmptyToNull() = $bEmptyIsNull) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTextBox.Control.InputRequired = $bInputRequired
		$iError = ($oTextBox.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bFilter <> Null) Then
		If Not IsBool($bFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTextBox.Control.UseFilterValueProposal = $bFilter
		$iError = ($oTextBox.Control.UseFilterValueProposal() = $bFilter) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTextBoxData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTextBoxGeneral
; Description ...: Set or Retrieve general Textbox control properties.
; Syntax ........: _LOWriter_FormConTextBoxGeneral(ByRef $oTextBox[, $sName = Null[, $oLabelField = Null[, $iTxtDir = Null[, $iMaxLen = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $bTabStop = Null[, $iTabOrder = Null[, $sDefaultTxt = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $iTextType = Null[, $bEndWithLF = Null[, $iScrollbars = Null[, $iPassChar = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oTextBox            - [in/out] an object. A Textbox Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control Name.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iMaxLen             - [optional] an integer value (-1-2147483647). Default is Null. The max length of text that can be entered. 0 = unlimited.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $sDefaultTxt         - [optional] a string value. Default is Null. The default text to display in the control.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iTextType           - [optional] an integer value (0-2). Default is Null. The text type. See Constants $LOW_FORM_CON_TXT_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bEndWithLF          - [optional] a boolean value. Default is Null. If True, the line ends will be a Line-Feed type, else a Carriage-Return plus a Line-Feed.
;                  $iScrollbars         - [optional] an integer value (0-3). Default is Null. The Scrollbars to use, if any. See Constants $LOW_FORM_CON_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iPassChar           - [optional] an integer value. Default is Null. The ASCII value of the Password character to display.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextBox not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTextBox not a Text Box Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 5 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 6 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $iMaxLen not an Integer, less than -1 or greater than 2147483647.
;                  @Error 1 @Extended 8 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 13 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 14 Return 0 = $sDefaultTxt not a String.
;                  @Error 1 @Extended 15 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 16 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 17 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 18 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 19 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 20 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 21 Return 0 = $iTextType not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 22 Return 0 = $bEndWithLF not a Boolean.
;                  @Error 1 @Extended 23 Return 0 = $iScrollbars not an Integer, less than 0 or greater than 3. See Constants $LOW_FORM_CON_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 24 Return 0 = $iPassChar not an Integer.
;                  @Error 1 @Extended 25 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 26 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 27 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 28 Return 0 = $sHelpURL not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $oLabelField
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $iMaxLen
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bVisible
;                  |                               64 = Error setting $bReadOnly
;                  |                               128 = Error setting $bPrintable
;                  |                               256 = Error setting $bTabStop
;                  |                               512 = Error setting $iTabOrder
;                  |                               1024 = Error setting $sDefaultTxt
;                  |                               2048 = Error setting $mFont
;                  |                               4096 = Error setting $iAlign
;                  |                               8192 = Error setting $iVertAlign
;                  |                               16384 = Error setting $iBackColor
;                  |                               32768 = Error setting $iBorder
;                  |                               65536 = Error setting $iBorderColor
;                  |                               131072 = Error setting $iTextType
;                  |                               262144 = Error setting $bEndWithLF
;                  |                               524288 = Error setting $iScrollbars
;                  |                               1048576 = Error setting $iPassChar
;                  |                               2097152 = Error setting $bHideSel
;                  |                               4194304 = Error setting $sAddInfo
;                  |                               8388608 = Error setting $sHelpText
;                  |                               16777216 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 25 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $sDefaultTxt, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormConTextBoxData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTextBoxGeneral(ByRef $oTextBox, $sName = Null, $oLabelField = Null, $iTxtDir = Null, $iMaxLen = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $bTabStop = Null, $iTabOrder = Null, $sDefaultTxt = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $iTextType = Null, $bEndWithLF = Null, $iScrollbars = Null, $iPassChar = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[25]
	Local Const $__LOW_FORM_CONTROL_LINE_END_CR = 0, $__LOW_FORM_CONTROL_LINE_END_LF = 1, $__LOW_FORM_CONTROL_LINE_END_CRLF = 2 ; "com.sun.star.awt.LineEndFormat"
	#forceref $__LOW_FORM_CONTROL_LINE_END_CR

	If Not IsObj($oTextBox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTextBox) <> $LOW_FORM_CON_TYPE_TEXT_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $oLabelField, $iTxtDir, $iMaxLen, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $bTabStop, $iTabOrder, $sDefaultTxt, $mFont, $iAlign, $iVertAlign, $iBackColor, $iBorder, $iBorderColor, $iTextType, $bEndWithLF, $iScrollbars, $iPassChar, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		__LO_ArrayFill($avControl, $oTextBox.Control.Name(), __LOWriter_FormConGetObj($oTextBox.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oTextBox.Control.WritingMode(), $oTextBox.Control.MaxTextLen(), _
				$oTextBox.Control.Enabled(), $oTextBox.Control.EnableVisible(), $oTextBox.Control.ReadOnly(), $oTextBox.Control.Printable(), $oTextBox.Control.Tabstop(), _
				$oTextBox.Control.TabIndex(), $oTextBox.Control.DefaultText(), __LOWriter_FormConSetGetFontDesc($oTextBox), $oTextBox.Control.Align(), $oTextBox.Control.VerticalAlign(), _
				$oTextBox.Control.BackgroundColor(), $oTextBox.Control.Border(), $oTextBox.Control.BorderColor(), _
				((($oTextBox.Control.MultiLine() = False) And ($oTextBox.Control.RichText() = False)) ? ($LOW_FORM_CON_TXT_TYPE_SINGLE_LINE) : (($oTextBox.Control.MultiLine() = True) And ($oTextBox.Control.RichText() = False)) ? ($LOW_FORM_CON_TXT_TYPE_MULTI_LINE) : ($LOW_FORM_CON_TXT_TYPE_MULTI_LINE_FORMATTED)), _ ; TextType setting.
				(($oTextBox.Control.LineEndFormat() = $__LOW_FORM_CONTROL_LINE_END_LF) ? (True) : (False)), _ ; Line Ending format
				((($oTextBox.Control.HScroll() = False) And ($oTextBox.Control.VScroll() = False)) ? ($LOW_FORM_CON_SCROLL_NONE) : _ ; Scrollbar mode.
				(($oTextBox.Control.HScroll() = True) And ($oTextBox.Control.VScroll() = False)) ? ($LOW_FORM_CON_SCROLL_HORI) : _ ; Scrollbar mode.
				((($oTextBox.Control.HScroll() = False) And ($oTextBox.Control.VScroll() = True)) ? ($LOW_FORM_CON_SCROLL_VERT) : ($LOW_FORM_CON_SCROLL_BOTH))), _ ; Scrollbar mode.
				$oTextBox.Control.EchoChar(), $oTextBox.Control.HideInactiveSelection(), $oTextBox.Control.Tag(), $oTextBox.Control.HelpText(), $oTextBox.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTextBox.Control.Name = $sName
		$iError = ($oTextBox.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oLabelField = Default) Then
		$oTextBox.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTextBox.Control.LabelControl = $oLabelField.Control()
		$iError = ($oTextBox.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oTextBox.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTextBox.Control.WritingMode = $iTxtDir
		$iError = ($oTextBox.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iMaxLen = Default) Then
		$oTextBox.Control.setPropertyToDefault("MaxTextLen")

	ElseIf ($iMaxLen <> Null) Then
		If Not __LO_IntIsBetween($iMaxLen, -1, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oTextBox.Control.MaxTextLen = $iMaxLen
		$iError = ($oTextBox.Control.MaxTextLen() = $iMaxLen) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oTextBox.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oTextBox.Control.Enabled = $bEnabled
		$iError = ($oTextBox.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bVisible = Default) Then
		$oTextBox.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oTextBox.Control.EnableVisible = $bVisible
		$iError = ($oTextBox.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bReadOnly = Default) Then
		$oTextBox.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oTextBox.Control.ReadOnly = $bReadOnly
		$iError = ($oTextBox.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bPrintable = Default) Then
		$oTextBox.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oTextBox.Control.Printable = $bPrintable
		$iError = ($oTextBox.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bTabStop = Default) Then
		$oTextBox.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oTextBox.Control.Tabstop = $bTabStop
		$iError = ($oTextBox.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 512) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oTextBox.Control.TabIndex = $iTabOrder
		$iError = ($oTextBox.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($sDefaultTxt = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default DefaultText.

	ElseIf ($sDefaultTxt <> Null) Then
		If Not IsString($sDefaultTxt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oTextBox.Control.DefaultText = $sDefaultTxt
		$iError = ($oTextBox.Control.DefaultText() = $sDefaultTxt) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 2048) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		__LOWriter_FormConSetGetFontDesc($oTextBox, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($iAlign = Default) Then
		$oTextBox.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oTextBox.Control.Align = $iAlign
		$iError = ($oTextBox.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iVertAlign = Default) Then
		$oTextBox.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oTextBox.Control.VerticalAlign = $iVertAlign
		$iError = ($oTextBox.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($iBackColor = Default) Then
		$oTextBox.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$oTextBox.Control.BackgroundColor = $iBackColor
		$iError = ($oTextBox.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($iBorder = Default) Then
		$oTextBox.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oTextBox.Control.Border = $iBorder
		$iError = ($oTextBox.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($iBorderColor = Default) Then
		$oTextBox.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oTextBox.Control.BorderColor = $iBorderColor
		$iError = ($oTextBox.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($iTextType = Default) Then
		$oTextBox.Control.setPropertyToDefault("MultiLine")
		$oTextBox.Control.setPropertyToDefault("RichText")

	ElseIf ($iTextType <> Null) Then
		If Not __LO_IntIsBetween($iTextType, $LOW_FORM_CON_TXT_TYPE_SINGLE_LINE, $LOW_FORM_CON_TXT_TYPE_MULTI_LINE_FORMATTED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		Switch $iTextType
			Case $LOW_FORM_CON_TXT_TYPE_SINGLE_LINE
				$oTextBox.Control.MultiLine = False
				$oTextBox.Control.RichText = False

				$iError = (($oTextBox.Control.MultiLine = False) And ($oTextBox.Control.RichText = False)) ? ($iError) : (BitOR($iError, 131072))

			Case $LOW_FORM_CON_TXT_TYPE_MULTI_LINE
				$oTextBox.Control.MultiLine = True
				$oTextBox.Control.RichText = False

				$iError = (($oTextBox.Control.MultiLine = True) And ($oTextBox.Control.RichText = False)) ? ($iError) : (BitOR($iError, 131072))

			Case $LOW_FORM_CON_TXT_TYPE_MULTI_LINE_FORMATTED
				$oTextBox.Control.MultiLine = True
				$oTextBox.Control.RichText = True

				$iError = (($oTextBox.Control.MultiLine = True) And ($oTextBox.Control.RichText = True)) ? ($iError) : (BitOR($iError, 131072))
		EndSwitch
	EndIf

	If ($bEndWithLF = Default) Then
		$oTextBox.Control.setPropertyToDefault("LineEndFormat")

	ElseIf ($bEndWithLF <> Null) Then
		If Not IsBool($bEndWithLF) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		If $bEndWithLF Then
			$oTextBox.Control.LineEndFormat = $__LOW_FORM_CONTROL_LINE_END_LF
			$iError = ($oTextBox.Control.LineEndFormat = $__LOW_FORM_CONTROL_LINE_END_LF) ? ($iError) : (BitOR($iError, 262144))

		Else
			$oTextBox.Control.LineEndFormat = $__LOW_FORM_CONTROL_LINE_END_CRLF
			$iError = ($oTextBox.Control.LineEndFormat = $__LOW_FORM_CONTROL_LINE_END_CRLF) ? ($iError) : (BitOR($iError, 262144))
		EndIf
	EndIf

	If ($iScrollbars = Default) Then
		$oTextBox.Control.setPropertyToDefault("HScroll")
		$oTextBox.Control.setPropertyToDefault("VScroll")

	ElseIf ($iScrollbars <> Null) Then
		If Not __LO_IntIsBetween($iScrollbars, $LOW_FORM_CON_SCROLL_NONE, $LOW_FORM_CON_SCROLL_BOTH) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		Switch $iScrollbars
			Case $LOW_FORM_CON_SCROLL_NONE
				$oTextBox.Control.HScroll = False
				$oTextBox.Control.VScroll = False
				$iError = (($oTextBox.Control.HScroll = False) And ($oTextBox.Control.VScroll = False)) ? ($iError) : (BitOR($iError, 524288))

			Case $LOW_FORM_CON_SCROLL_HORI
				$oTextBox.Control.HScroll = True
				$oTextBox.Control.VScroll = False
				$iError = (($oTextBox.Control.HScroll = True) And ($oTextBox.Control.VScroll = False)) ? ($iError) : (BitOR($iError, 524288))

			Case $LOW_FORM_CON_SCROLL_VERT
				$oTextBox.Control.HScroll = False
				$oTextBox.Control.VScroll = True
				$iError = (($oTextBox.Control.HScroll = False) And ($oTextBox.Control.VScroll = True)) ? ($iError) : (BitOR($iError, 524288))

			Case $LOW_FORM_CON_SCROLL_BOTH
				$oTextBox.Control.HScroll = True
				$oTextBox.Control.VScroll = True
				$iError = (($oTextBox.Control.HScroll = True) And ($oTextBox.Control.VScroll = True)) ? ($iError) : (BitOR($iError, 524288))
		EndSwitch
	EndIf

	If ($iPassChar = Default) Then
		$oTextBox.Control.setPropertyToDefault("EchoChar")

	ElseIf ($iPassChar <> Null) Then
		If Not IsInt($iPassChar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oTextBox.Control.EchoChar = $iPassChar
		$iError = ($oTextBox.Control.EchoChar() = $iPassChar) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($bHideSel = Default) Then
		$oTextBox.Control.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oTextBox.Control.HideInactiveSelection = $bHideSel
		$iError = ($oTextBox.Control.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 4194304) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, 0)

		$oTextBox.Control.Tag = $sAddInfo
		$iError = ($oTextBox.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	If ($sHelpText = Default) Then
		$oTextBox.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 27, 0)

		$oTextBox.Control.HelpText = $sHelpText
		$iError = ($oTextBox.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 8388608))
	EndIf

	If ($sHelpURL = Default) Then
		$oTextBox.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 28, 0)

		$oTextBox.Control.HelpURL = $sHelpURL
		$iError = ($oTextBox.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 16777216))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTextBoxGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTimeFieldData
; Description ...: Set or Retrieve Time Field Data Properties.
; Syntax ........: _LOWriter_FormConTimeFieldData(ByRef $oTimeField[, $sDataField = Null[, $bInputRequired = Null]])
; Parameters ....: $oTimeField          - [in/out] an object. A Time Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The Datafield name to retrieve content from, either a Table name, SQL query, or other.
;                  $bInputRequired      - [optional] a boolean value. Default is Null. If True, the control requires input.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTimeField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTimeField not a Time Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bInputRequired not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  |                               2 = Error setting $bInputRequired
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
; Related .......: _LOWriter_FormConTimeFieldValue, _LOWriter_FormConTimeFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTimeFieldData(ByRef $oTimeField, $sDataField = Null, $bInputRequired = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[2]

	If Not IsObj($oTimeField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTimeField) <> $LOW_FORM_CON_TYPE_TIME_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField, $bInputRequired) Then
		__LO_ArrayFill($avControl, $oTimeField.Control.DataField(), $oTimeField.Control.InputRequired())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sDataField <> Null) Then
		If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTimeField.Control.DataField = $sDataField
		$iError = ($oTimeField.Control.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bInputRequired <> Null) Then
		If Not IsBool($bInputRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oTimeField.Control.InputRequired = $bInputRequired
		$iError = ($oTimeField.Control.InputRequired() = $bInputRequired) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTimeFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTimeFieldGeneral
; Description ...: Set or Retrieve general Time Field properties.
; Syntax ........: _LOWriter_FormConTimeFieldGeneral(ByRef $oTimeField[, $sName = Null[, $oLabelField = Null[, $iTxtDir = Null[, $bStrict = Null[, $bEnabled = Null[, $bVisible = Null[, $bReadOnly = Null[, $bPrintable = Null[, $iMouseScroll = Null[, $bTabStop = Null[, $iTabOrder = Null[, $tTimeMin = Null[, $tTimeMax = Null[, $iFormat = Null[, $tTimeDefault = Null[, $bSpin = Null[, $bRepeat = Null[, $iDelay = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iBackColor = Null[, $iBorder = Null[, $iBorderColor = Null[, $bHideSel = Null[, $sAddInfo = Null[, $sHelpText = Null[, $sHelpURL = Null]]]]]]]]]]]]]]]]]]]]]]]]]]]])
; Parameters ....: $oTimeField          - [in/out] an object.object. A Time Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $oLabelField         - [optional] an object. Default is Null. A Label Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $iTxtDir             - [optional] an integer value (0-5). Default is Null. The Text direction. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bStrict             - [optional] a boolean value. Default is Null. If True, strict formatting is enabled.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the control is enabled.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the control is visible.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, the control is Read-Only.
;                  $bPrintable          - [optional] a boolean value. Default is Null. If True, the control will be displayed when printed.
;                  $iMouseScroll        - [optional] an integer value (0-2). Default is Null. The behavior of the mouse scroll wheel on the Control. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTabStop            - [optional] a boolean value. Default is Null. If True, the control can be selected with the Tab key.
;                  $iTabOrder           - [optional] an integer value (0-2147483647). Default is Null. The order the control is focused by the Tab key.
;                  $tTimeMin            - [optional] a dll struct value. Default is Null. The minimum time	 allowed to be entered, created previously by _LOWriter_DateStructCreate.
;                  $tTimeMax            - [optional] a dll struct value. Default is Null. The maximum time	 allowed to be entered, created previously by _LOWriter_DateStructCreate.
;                  $iFormat             - [optional] an integer value (0-5). Default is Null. The Time Format to display the content in. See Constants $LOW_FORM_CON_TIME_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $tTimeDefault        - [optional] a dll struct value. Default is Null. The Default time to display, created previously by _LOWriter_DateStructCreate.
;                  $bSpin               - [optional] a boolean value. Default is Null. If True, the field will act as a spin button.
;                  $bRepeat             - [optional] a boolean value. Default is Null. If True, the button action will repeat if the button is clicked and held down.
;                  $iDelay              - [optional] an integer value (0-2147483647). Default is Null. The delay between button repeats, set in milliseconds.
;                  $mFont               - [optional] a map. Default is Null. The Font descriptor to use. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBackColor          - [optional] an integer value (0-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBorder             - [optional] an integer value (0-2). Default is Null. The Border Style. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iBorderColor        - [optional] an integer value (0-16777215). Default is Null. The Border color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bHideSel            - [optional] a boolean value. Default is Null. If True, any selections in the control will not remain selected when the control loses focus.
;                  $sAddInfo            - [optional] a string value. Default is Null. Additional information text.
;                  $sHelpText           - [optional] a string value. Default is Null. The Help text.
;                  $sHelpURL            - [optional] a string value. Default is Null. The Help URL.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTimeField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTimeField not a Time Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $oLabelField not an Object.
;                  @Error 1 @Extended 5 Return 0 = Object called in $oLabelField not a Label Control.
;                  @Error 1 @Extended 6 Return 0 = $iTxtDir not an Integer, less than 0 or greater than 5. See Constants $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bStrict not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $bPrintable not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iMouseScroll not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_MOUSE_SCROLL_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $bTabStop not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $iTabOrder not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 15 Return 0 = $tTimeMin not an Object.
;                  @Error 1 @Extended 16 Return 0 = $tTimeMax not an Object.
;                  @Error 1 @Extended 17 Return 0 = $iFormat not an Integer, less than 0 or greater than 5. See Constants $LOW_FORM_CON_TIME_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 18 Return 0 = $tTimeDefault not an Object.
;                  @Error 1 @Extended 19 Return 0 = $bSpin not a Boolean.
;                  @Error 1 @Extended 20 Return 0 = $bRepeat not a Boolean.
;                  @Error 1 @Extended 21 Return 0 = $iDelay not an Integer, less than 0 or greater than 2147483647.
;                  @Error 1 @Extended 22 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 23 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 24 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOW_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 25 Return 0 = $iBackColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 26 Return 0 = $iBorder not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CON_BORDER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 27 Return 0 = $iBorderColor not an Integer, less than 0 or greater than 16777215.
;                  @Error 1 @Extended 28 Return 0 = $bHideSel not a Boolean.
;                  @Error 1 @Extended 29 Return 0 = $sAddInfo not a String.
;                  @Error 1 @Extended 30 Return 0 = $sHelpText not a String.
;                  @Error 1 @Extended 31 Return 0 = $sHelpURL not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.DateTime" Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.util.Time" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve current Minimum Time.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve current Maximum Time.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $oLabelField
;                  |                               4 = Error setting $iTxtDir
;                  |                               8 = Error setting $bStrict
;                  |                               16 = Error setting $bEnabled
;                  |                               32 = Error setting $bVisible
;                  |                               64 = Error setting $bReadOnly
;                  |                               128 = Error setting $bPrintable
;                  |                               256 = Error setting $iMouseScroll
;                  |                               512 = Error setting $bTabStop
;                  |                               1024 = Error setting $iTabOrder
;                  |                               2048 = Error setting $tTimeMin
;                  |                               4096 = Error setting $tTimeMax
;                  |                               8192 = Error setting $iFormat
;                  |                               16384 = Error setting $tTimeDefault
;                  |                               32768 = Error setting $bSpin
;                  |                               65536 = Error setting $bRepeat
;                  |                               131072 = Error setting $iDelay
;                  |                               262144 = Error setting $mFont
;                  |                               524288 = Error setting $iAlign
;                  |                               1048576 = Error setting $iVertAlign
;                  |                               2097152 = Error setting $iBackColor
;                  |                               4194304 = Error setting $iBorder
;                  |                               8388608 = Error setting $iBorderColor
;                  |                               16777216 = Error setting $bHideSel
;                  |                               33554432 = Error setting $sAddInfo
;                  |                               67108864 = Error setting $sHelpText
;                  |                               134217728 = Error setting $sHelpURL
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 28 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Call any optional parameter with Default keyword to reset the value to default. This can include a default of "Null", or "Default", etc., that is otherwise impossible to set.
;                  Some parameters cannot be returned to default using the Default keyword, namely: $sName, $iTabOrder, $mFont, $sAddInfo.
;                  Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, https://bugs.documentfoundation.org/show_bug.cgi?id=131196
; Related .......: _LOWriter_FormatKeyCreate, _LOWriter_FormatKeysGetList, _LOWriter_FormConTimeFieldValue, _LOWriter_FormConTimeFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTimeFieldGeneral(ByRef $oTimeField, $sName = Null, $oLabelField = Null, $iTxtDir = Null, $bStrict = Null, $bEnabled = Null, $bVisible = Null, $bReadOnly = Null, $bPrintable = Null, $iMouseScroll = Null, $bTabStop = Null, $iTabOrder = Null, $tTimeMin = Null, $tTimeMax = Null, $iFormat = Null, $tTimeDefault = Null, $bSpin = Null, $bRepeat = Null, $iDelay = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iBackColor = Null, $iBorder = Null, $iBorderColor = Null, $bHideSel = Null, $sAddInfo = Null, $sHelpText = Null, $sHelpURL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tTime, $tCurMin, $tCurMax, $tCurDefault
	Local $avControl[28]

	If Not IsObj($oTimeField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTimeField) <> $LOW_FORM_CON_TYPE_TIME_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $oLabelField, $iTxtDir, $bStrict, $bEnabled, $bVisible, $bReadOnly, $bPrintable, $iMouseScroll, $bTabStop, $iTabOrder, $tTimeMin, $tTimeMax, $iFormat, $tTimeDefault, $bSpin, $bRepeat, $iDelay, $mFont, $iAlign, $iVertAlign, $iBackColor, $iBorder, $iBorderColor, $bHideSel, $sAddInfo, $sHelpText, $sHelpURL) Then
		$tTime = $oTimeField.Control.TimeMin()
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$tCurMin = __LO_CreateStruct("com.sun.star.util.DateTime")
		If Not IsObj($tCurMin) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tCurMin.Hours = $tTime.Hours()
		$tCurMin.Minutes = $tTime.Minutes()
		$tCurMin.Seconds = $tTime.Seconds()
		$tCurMin.NanoSeconds = $tTime.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tCurMin.IsUTC = $tTime.IsUTC()

		$tTime = $oTimeField.Control.TimeMax()
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$tCurMax = __LO_CreateStruct("com.sun.star.util.DateTime")
		If Not IsObj($tCurMax) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tCurMax.Hours = $tTime.Hours()
		$tCurMax.Minutes = $tTime.Minutes()
		$tCurMax.Seconds = $tTime.Seconds()
		$tCurMax.NanoSeconds = $tTime.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tCurMax.IsUTC = $tTime.IsUTC()

		$tTime = $oTimeField.Control.DefaultTime() ; Default time is Null when not set.
		If IsObj($tTime) Then
			$tCurDefault = __LO_CreateStruct("com.sun.star.util.DateTime")
			If Not IsObj($tCurDefault) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tCurDefault.Hours = $tTime.Hours()
			$tCurDefault.Minutes = $tTime.Minutes()
			$tCurDefault.Seconds = $tTime.Seconds()
			$tCurDefault.NanoSeconds = $tTime.NanoSeconds()
			If __LO_VersionCheck(4.1) Then $tCurDefault.IsUTC = $tTime.IsUTC()

		Else
			$tCurDefault = $tTime
		EndIf

		__LO_ArrayFill($avControl, $oTimeField.Control.Name(), __LOWriter_FormConGetObj($oTimeField.Control.LabelControl(), $LOW_FORM_CON_TYPE_LABEL), $oTimeField.Control.WritingMode(), $oTimeField.Control.StrictFormat(), _
				$oTimeField.Control.Enabled(), $oTimeField.Control.EnableVisible(), $oTimeField.Control.ReadOnly(), $oTimeField.Control.Printable(), $oTimeField.Control.MouseWheelBehavior(), _
				$oTimeField.Control.Tabstop(), $oTimeField.Control.TabIndex(), $tCurMin, $tCurMax, $oTimeField.Control.TimeFormat(), $tCurDefault, $oTimeField.Control.Spin(), _
				$oTimeField.Control.Repeat(), $oTimeField.Control.RepeatDelay(), __LOWriter_FormConSetGetFontDesc($oTimeField), $oTimeField.Control.Align(), $oTimeField.Control.VerticalAlign(), _
				$oTimeField.Control.BackgroundColor(), $oTimeField.Control.Border(), $oTimeField.Control.BorderColor(), $oTimeField.Control.HideInactiveSelection(), _
				$oTimeField.Control.Tag(), $oTimeField.Control.HelpText(), $oTimeField.Control.HelpURL())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName = Default) Then
		$iError = BitOR($iError, 1) ; Can't Default Name.

	ElseIf ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oTimeField.Control.Name = $sName
		$iError = ($oTimeField.Control.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oLabelField = Default) Then
		$oTimeField.Control.setPropertyToDefault("LabelControl")

	ElseIf ($oLabelField <> Null) Then
		If Not IsObj($oLabelField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If (__LOWriter_FormConIdentify($oLabelField) <> $LOW_FORM_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oTimeField.Control.LabelControl = $oLabelField.Control()
		$iError = ($oTimeField.Control.LabelControl() = $oLabelField.Control()) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTxtDir = Default) Then
		$oTimeField.Control.setPropertyToDefault("WritingMode")

	ElseIf ($iTxtDir <> Null) Then
		If Not __LO_IntIsBetween($iTxtDir, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oTimeField.Control.WritingMode = $iTxtDir
		$iError = ($oTimeField.Control.WritingMode() = $iTxtDir) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bStrict = Default) Then
		$oTimeField.Control.setPropertyToDefault("StrictFormat")

	ElseIf ($bStrict <> Null) Then
		If Not IsBool($bStrict) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oTimeField.Control.StrictFormat = $bStrict
		$iError = ($oTimeField.Control.StrictFormat() = $bStrict) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEnabled = Default) Then
		$oTimeField.Control.setPropertyToDefault("Enabled")

	ElseIf ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oTimeField.Control.Enabled = $bEnabled
		$iError = ($oTimeField.Control.Enabled() = $bEnabled) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bVisible = Default) Then
		$oTimeField.Control.setPropertyToDefault("EnableVisible")

	ElseIf ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oTimeField.Control.EnableVisible = $bVisible
		$iError = ($oTimeField.Control.EnableVisible() = $bVisible) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bReadOnly = Default) Then
		$oTimeField.Control.setPropertyToDefault("ReadOnly")

	ElseIf ($bReadOnly <> Null) Then
		If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oTimeField.Control.ReadOnly = $bReadOnly
		$iError = ($oTimeField.Control.ReadOnly() = $bReadOnly) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bPrintable = Default) Then
		$oTimeField.Control.setPropertyToDefault("Printable")

	ElseIf ($bPrintable <> Null) Then
		If Not IsBool($bPrintable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oTimeField.Control.Printable = $bPrintable
		$iError = ($oTimeField.Control.Printable() = $bPrintable) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iMouseScroll = Default) Then
		$oTimeField.Control.setPropertyToDefault("MouseWheelBehavior")

	ElseIf ($iMouseScroll <> Null) Then
		If Not __LO_IntIsBetween($iMouseScroll, $LOW_FORM_CON_MOUSE_SCROLL_DISABLED, $LOW_FORM_CON_MOUSE_SCROLL_ALWAYS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oTimeField.Control.MouseWheelBehavior = $iMouseScroll
		$iError = ($oTimeField.Control.MouseWheelBehavior() = $iMouseScroll) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bTabStop = Default) Then
		$oTimeField.Control.setPropertyToDefault("Tabstop")

	ElseIf ($bTabStop <> Null) Then
		If Not IsBool($bTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oTimeField.Control.Tabstop = $bTabStop
		$iError = ($oTimeField.Control.Tabstop() = $bTabStop) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($iTabOrder = Default) Then
		$iError = BitOR($iError, 1024) ; Can't Default TabIndex.

	ElseIf ($iTabOrder <> Null) Then
		If Not __LO_IntIsBetween($iTabOrder, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oTimeField.Control.TabIndex = $iTabOrder
		$iError = ($oTimeField.Control.TabIndex() = $iTabOrder) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($tTimeMin = Default) Then
		$oTimeField.Control.setPropertyToDefault("TimeMin")

	ElseIf ($tTimeMin <> Null) Then
		If Not IsObj($tTimeMin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$tTime = __LO_CreateStruct("com.sun.star.util.Time")
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tTime.Hours = $tTimeMin.Hours()
		$tTime.Minutes = $tTimeMin.Minutes()
		$tTime.Seconds = $tTimeMin.Seconds()
		$tTime.NanoSeconds = $tTimeMin.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tTime.IsUTC = $tTimeMin.IsUTC()

		$oTimeField.Control.TimeMin = $tTime
		$iError = (__LOWriter_DateStructCompare($oTimeField.Control.TimeMin(), $tTime, False, True)) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($tTimeMax = Default) Then
		$oTimeField.Control.setPropertyToDefault("TimeMax")

	ElseIf ($tTimeMax <> Null) Then
		If Not IsObj($tTimeMax) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$tTime = __LO_CreateStruct("com.sun.star.util.Time")
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tTime.Hours = $tTimeMax.Hours()
		$tTime.Minutes = $tTimeMax.Minutes()
		$tTime.Seconds = $tTimeMax.Seconds()
		$tTime.NanoSeconds = $tTimeMax.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tTime.IsUTC = $tTimeMax.IsUTC()

		$oTimeField.Control.TimeMax = $tTime
		$iError = (__LOWriter_DateStructCompare($oTimeField.Control.TimeMax(), $tTime, False, True)) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iFormat = Default) Then
		$oTimeField.Control.setPropertyToDefault("TimeFormat")

	ElseIf ($iFormat <> Null) Then
		If Not __LO_IntIsBetween($iFormat, $LOW_FORM_CON_TIME_FRMT_24_SHORT, $LOW_FORM_CON_TIME_FRMT_DURATION_LONG) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oTimeField.Control.TimeFormat = $iFormat
		$iError = ($oTimeField.Control.TimeFormat() = $iFormat) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	If ($tTimeDefault = Default) Then
		$oTimeField.Control.setPropertyToDefault("DefaultTime")

	ElseIf ($tTimeDefault <> Null) Then
		If Not IsObj($tTimeDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)

		$tTime = __LO_CreateStruct("com.sun.star.util.Time")
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tTime.Hours = $tTimeDefault.Hours()
		$tTime.Minutes = $tTimeDefault.Minutes()
		$tTime.Seconds = $tTimeDefault.Seconds()
		$tTime.NanoSeconds = $tTimeDefault.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tTime.IsUTC = $tTimeDefault.IsUTC()

		$oTimeField.Control.DefaultTime = $tTime
		$iError = (__LOWriter_DateStructCompare($oTimeField.Control.DefaultTime(), $tTime, False, True)) ? ($iError) : (BitOR($iError, 16384))
	EndIf

	If ($bSpin = Default) Then
		$oTimeField.Control.setPropertyToDefault("Spin")

	ElseIf ($bSpin <> Null) Then
		If Not IsBool($bSpin) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

		$oTimeField.Control.Spin = $bSpin
		$iError = ($oTimeField.Control.Spin() = $bSpin) ? ($iError) : (BitOR($iError, 32768))
	EndIf

	If ($bRepeat = Default) Then
		$oTimeField.Control.setPropertyToDefault("Repeat")

	ElseIf ($bRepeat <> Null) Then
		If Not IsBool($bRepeat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 20, 0)

		$oTimeField.Control.Repeat = $bRepeat
		$iError = ($oTimeField.Control.Repeat() = $bRepeat) ? ($iError) : (BitOR($iError, 65536))
	EndIf

	If ($iDelay = Default) Then
		$oTimeField.Control.setPropertyToDefault("RepeatDelay")

	ElseIf ($iDelay <> Null) Then
		If Not __LO_IntIsBetween($iDelay, 0, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 21, 0)

		$oTimeField.Control.RepeatDelay = $iDelay
		$iError = ($oTimeField.Control.RepeatDelay() = $iDelay) ? ($iError) : (BitOR($iError, 131072))
	EndIf

	If ($mFont = Default) Then
		$iError = BitOR($iError, 262144) ; Can't Default Font (Works, but doesn't change the font).

	ElseIf ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 22, 0)

		__LOWriter_FormConSetGetFontDesc($oTimeField, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 262144))
	EndIf

	If ($iAlign = Default) Then
		$oTimeField.Control.setPropertyToDefault("Align")

	ElseIf ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOW_ALIGN_HORI_LEFT, $LOW_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 23, 0)

		$oTimeField.Control.Align = $iAlign
		$iError = ($oTimeField.Control.Align() = $iAlign) ? ($iError) : (BitOR($iError, 524288))
	EndIf

	If ($iVertAlign = Default) Then
		$oTimeField.Control.setPropertyToDefault("VerticalAlign")

	ElseIf ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_ALIGN_VERT_TOP, $LOW_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 24, 0)

		$oTimeField.Control.VerticalAlign = $iVertAlign
		$iError = ($oTimeField.Control.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 1048576))
	EndIf

	If ($iBackColor = Default) Then
		$oTimeField.Control.setPropertyToDefault("BackgroundColor")

	ElseIf ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 25, 0)

		$oTimeField.Control.BackgroundColor = $iBackColor
		$iError = ($oTimeField.Control.BackgroundColor() = $iBackColor) ? ($iError) : (BitOR($iError, 2097152))
	EndIf

	If ($iBorder = Default) Then
		$oTimeField.Control.setPropertyToDefault("Border")

	ElseIf ($iBorder <> Null) Then
		If Not __LO_IntIsBetween($iBorder, $LOW_FORM_CON_BORDER_WITHOUT, $LOW_FORM_CON_BORDER_FLAT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 26, 0)

		$oTimeField.Control.Border = $iBorder
		$iError = ($oTimeField.Control.Border() = $iBorder) ? ($iError) : (BitOR($iError, 4194304))
	EndIf

	If ($iBorderColor = Default) Then
		$oTimeField.Control.setPropertyToDefault("BorderColor")

	ElseIf ($iBorderColor <> Null) Then
		If Not __LO_IntIsBetween($iBorderColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 27, 0)

		$oTimeField.Control.BorderColor = $iBorderColor
		$iError = ($oTimeField.Control.BorderColor() = $iBorderColor) ? ($iError) : (BitOR($iError, 8388608))
	EndIf

	If ($bHideSel = Default) Then
		$oTimeField.Control.setPropertyToDefault("HideInactiveSelection")

	ElseIf ($bHideSel <> Null) Then
		If Not IsBool($bHideSel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 28, 0)

		$oTimeField.Control.HideInactiveSelection = $bHideSel
		$iError = ($oTimeField.Control.HideInactiveSelection() = $bHideSel) ? ($iError) : (BitOR($iError, 16777216))
	EndIf

	If ($sAddInfo = Default) Then
		$iError = BitOR($iError, 33554432) ; Can't Default Tag.

	ElseIf ($sAddInfo <> Null) Then
		If Not IsString($sAddInfo) Then Return SetError($__LO_STATUS_INPUT_ERROR, 29, 0)

		$oTimeField.Control.Tag = $sAddInfo
		$iError = ($oTimeField.Control.Tag() = $sAddInfo) ? ($iError) : (BitOR($iError, 33554432))
	EndIf

	If ($sHelpText = Default) Then
		$oTimeField.Control.setPropertyToDefault("HelpText")

	ElseIf ($sHelpText <> Null) Then
		If Not IsString($sHelpText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 30, 0)

		$oTimeField.Control.HelpText = $sHelpText
		$iError = ($oTimeField.Control.HelpText() = $sHelpText) ? ($iError) : (BitOR($iError, 67108864))
	EndIf

	If ($sHelpURL = Default) Then
		$oTimeField.Control.setPropertyToDefault("HelpURL")

	ElseIf ($sHelpURL <> Null) Then
		If Not IsString($sHelpURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 31, 0)

		$oTimeField.Control.HelpURL = $sHelpURL
		$iError = ($oTimeField.Control.HelpURL() = $sHelpURL) ? ($iError) : (BitOR($iError, 134217728))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTimeFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormConTimeFieldValue
; Description ...: Set or retrieve the current Time field value.
; Syntax ........: _LOWriter_FormConTimeFieldValue(ByRef $oTimeField[, $tTimeValue = Null])
; Parameters ....: $oTimeField          - [in/out] an object. A Time Field Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $tTimeValue          - [optional] a dll struct value. Default is Null. The time value to set the field to, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Structure
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTimeField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTimeField not a Time Field Control.
;                  @Error 1 @Extended 3 Return 0 = $tTimeValue not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.DateTime" Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.util.Time" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $tTimeValue
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Structure = Success. All optional parameters were called with Null, returning current Time value as a Time Structure.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current Time value. Return will be Null if the Time hasn't been set.
;                  Call $tTimeValue with Default keyword to reset the value to default.
; Related .......: _LOWriter_FormConTimeFieldGeneral, _LOWriter_FormConTimeFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormConTimeFieldValue(ByRef $oTimeField, $tTimeValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tTime, $tCurTime

	If Not IsObj($oTimeField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOWriter_FormConIdentify($oTimeField) <> $LOW_FORM_CON_TYPE_TIME_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($tTimeValue) Then
		$tTime = $oTimeField.Control.Time() ; Time is Null when not set.
		If IsObj($tTime) Then
			$tCurTime = __LO_CreateStruct("com.sun.star.util.DateTime")
			If Not IsObj($tCurTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tCurTime.Hours = $tTime.Hours()
			$tCurTime.Minutes = $tTime.Minutes()
			$tCurTime.Seconds = $tTime.Seconds()
			$tCurTime.NanoSeconds = $tTime.NanoSeconds()
			If __LO_VersionCheck(4.1) Then $tCurTime.IsUTC = $tTime.IsUTC()

		Else
			$tCurTime = $tTime
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $tCurTime)
	EndIf

	If ($tTime = Default) Then
		$oTimeField.Control.setPropertyToDefault("Time")

	Else
		If Not IsObj($tTimeValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tTime = __LO_CreateStruct("com.sun.star.util.Time")
		If Not IsObj($tTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tTime.Hours = $tTimeValue.Hours()
		$tTime.Minutes = $tTimeValue.Minutes()
		$tTime.Seconds = $tTimeValue.Seconds()
		$tTime.NanoSeconds = $tTimeValue.NanoSeconds()
		If __LO_VersionCheck(4.1) Then $tTime.IsUTC = $tTimeValue.IsUTC()

		$oTimeField.Control.Time = $tTime
		$iError = (__LOWriter_DateStructCompare($oTimeField.Control.Time(), $tTime, False, True)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormConTimeFieldValue

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormDelete
; Description ...: Delete a form or sub-form from a document.
; Syntax ........: _LOWriter_FormDelete(ByRef $oForm)
; Parameters ....: $oForm               - [in/out] an object. A Form Object returned from a previous _LOWriter_FormsGetList, or _LOWriter_FormAdd function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oForm not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oForm not a Form Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve parent Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve parent Document Object.
;                  @Error 3 @Extended 3 Return 0 = Parent Document is Read Only.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Forms Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to modify Form name.
;                  @Error 3 @Extended 6 Return 0 = Failed to delete form.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Form was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormAdd, _LOWriter_FormsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormDelete(ByRef $oForm)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oParent, $oDoc
	Local $sTempName = "AutoIt_FORM_"
	Local $iCount = 1

	If Not IsObj($oForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oForm.supportsService("com.sun.star.form.component.Form") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oParent = $oForm.getParent()
	If Not IsObj($oParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oDoc = $oForm ; Identify the parent document.

	Do
		$oDoc = $oDoc.getParent()
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Until $oDoc.supportsService("com.sun.star.text.TextDocument")

	If $oDoc.IsReadOnly() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If $oParent.supportsService("com.sun.star.text.TextDocument") Then
		$oParent = $oParent.DrawPage.Forms()
		If Not IsObj($oParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	; Create a Unique name.
	While $oParent.hasByName($sTempName & $iCount)
		$iCount += 1
	WEnd

	; Rename the form.
	$oForm.Name = $sTempName & $iCount
	If Not ($oForm.Name() = $sTempName & $iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$oParent.removeByName($sTempName & $iCount)

	If $oParent.hasByName($sTempName & $iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FormDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormGetObjByIndex
; Description ...: Retrieve a Form Object by index.
; Syntax ........: _LOWriter_FormGetObjByIndex(ByRef $oObj, $iForm)
; Parameters ....: $oObj                - [in/out] an object. Either a Document Object or a Form object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function, or a Form Object returned from a previous _LOWriter_FormsGetList, or _LOWriter_FormAdd function.
;                  $iForm               - an integer value. The Index value of the Form to retrieve. 0 Based.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iForm not an Integer, less than 0 or greater then number of forms contained in object.
;                  @Error 1 @Extended 3 Return 0 = Called Object in $oObj, not a Document Object, and not a Form Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of forms.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Form Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested form Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormsGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormGetObjByIndex(ByRef $oObj, $iForm)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iForms, $iCount = -1
	Local $oForm

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iForms = _LOWriter_FormsGetCount($oObj)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iForm, 0, ($iForms - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oObj.supportsService("com.sun.star.form.component.Form") Then
		For $i = 0 To $oObj.Count() - 1
			If $oObj.getByIndex($i).supportsService("com.sun.star.form.component.Form") Then $iCount += 1
			If ($iCount = $iForm) Then
				$oForm = $oObj.getByIndex($i)
				ExitLoop
			EndIf
		Next

	ElseIf $oObj.supportsService("com.sun.star.text.TextDocument") Then
		$oForm = $oObj.DrawPage.Forms.getByIndex($iForm)

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; wrong type of input item.
	EndIf

	If Not IsObj($oForm) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oForm)
EndFunc   ;==>_LOWriter_FormGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormParent
; Description ...: Set or Retrieve a Form's parent, i.e. set a form as Sub-Form, or a Sub-Form as a top-level form.
; Syntax ........: _LOWriter_FormParent(ByRef $oForm[, $oParent = Null])
; Parameters ....: $oForm               - [in/out] an object. A Form Object returned from a previous _LOWriter_FormsGetList, or _LOWriter_FormAdd function.
;                  $oParent             - [optional] an object. Default is Null. A Document or Form object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function, or _LOWriter_FormsGetList, or _LOWriter_FormAdd function.
; Return values .: Success: 1 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oForm not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oForm not a Form Object.
;                  @Error 1 @Extended 3 Return 0 = $oParent not an Object.
;                  @Error 1 @Extended 4 Return 0 = Object called in $oParent not an Form and not a Document.
;                  @Error 1 @Extended 5 Return 0 = Destination called in $oParent is the same as form's current parent.
;                  @Error 1 @Extended 6 Return 0 = Destination called in $oParent is the same as form called in $oForm.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to clone the form Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Form's parent Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Document's Forms Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Destination Parent Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Form's name.
;                  @Error 3 @Extended 5 Return 0 = Failed to rename Form.
;                  @Error 3 @Extended 6 Return 0 = Failed to insert cloned form into destination.
;                  @Error 3 @Extended 7 Return 0 = Failed to delete original form.
;                  @Error 3 @Extended 8 Return 0 = Failed to retrieve new form's Object.
;                  @Error 3 @Extended 9 Return 0 = Failed to set form's name back to original name.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Form was successfully moved, Object called in $oForm has been updated to new Object.
;                  @Error 0 @Extended 1 Return Object = Success. Returning Form's parent Object, which is a Document Object.
;                  @Error 0 @Extended 2 Return Object = Success. Returning Form's parent Object, which is a Form Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function allows you to change a sub-form into being a top-level form, change a top-level form into being a sub-form, or move a sub-form to be a sub-form of another form.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the Form's parent Object.
;                  If the parent Object is a Document, that means the Form is a top-level form. Otherwise it is a Sub-Form.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormParent(ByRef $oForm, $oParent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oOldParent, $oDestParent = $oParent, $oNewForm, $oFormCopy
	Local $sTempName = "AutoIt_FORM_", $sOldName
	Local $iCount = 1

	If Not IsObj($oForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oForm.supportsService("com.sun.star.form.component.Form") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oOldParent = $oForm.getParent()
	If Not IsObj($oOldParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($oParent) Then
		If $oOldParent.supportsService("com.sun.star.text.TextDocument") Then ; Parent is a Document.

			Return SetError($__LO_STATUS_SUCCESS, 1, $oOldParent)

		Else ; Parent is a Form.

			Return SetError($__LO_STATUS_SUCCESS, 2, $oOldParent)
		EndIf
	EndIf

	If Not IsObj($oParent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not $oParent.supportsService("com.sun.star.form.component.Form") And Not $oParent.supportsService("com.sun.star.text.TextDocument") Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($oForm.Parent() = $oParent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($oForm = $oParent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If $oOldParent.supportsService("com.sun.star.text.TextDocument") Then $oOldParent = $oOldParent.DrawPage.Forms()
	If Not IsObj($oOldParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not $oDestParent.supportsService("com.sun.star.form.component.Form") Then $oDestParent = $oParent.DrawPage.Forms() ; Destination is a document.
	If Not IsObj($oDestParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	; Create a Unique name.
	While ($oOldParent.hasByName($sTempName & $iCount) And $oDestParent.hasByName($sTempName & $iCount))
		$iCount += 1
	WEnd

	$sTempName = $sTempName & $iCount

	; Retrieve the Form's original name.
	$sOldName = $oForm.Name()
	If Not IsString($sOldName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	; Rename the form.
	$oForm.Name = $sTempName
	If ($oForm.Name() <> ($sTempName)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$oFormCopy = $oForm.CreateClone()
	If Not IsObj($oFormCopy) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDestParent.insertByName($sTempName, $oFormCopy)
	If Not $oDestParent.hasByName($sTempName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	$oOldParent.removeByName($sTempName)
	If $oOldParent.hasByName($sTempName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

	$oNewForm = $oDestParent.getByName($sTempName)
	If Not IsObj($oNewForm) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 8, 0)

	$oForm = $oNewForm

	$oForm.Name = $sOldName
	If ($oForm.Name() <> $sOldName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 9, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FormParent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormPropertiesData
; Description ...: Set or Retrieve Form Data properties.
; Syntax ........: _LOWriter_FormPropertiesData(ByRef $oForm[, $sSource = Null[, $iContentType = Null[, $sContent = Null[, $bAnalyzeSQL = Null[, $sFilter = Null[, $sSort = Null[, $aLinkMaster = Null[, $aLinkSlave = Null[, $bAdditions = Null[, $bModifications = Null[, $bDeletions = Null[, $bAddOnly = Null[, $iNavBar = Null[, $iCycle = Null]]]]]]]]]]]]]])
; Parameters ....: $oForm               - [in/out] an object. A Form object returned by a previous _LOWriter_FormAdd, _LOWriter_FormGetObjByIndex or _LOWriter_FormsGetList function.
;                  $sSource             - [optional] a string value. Default is Null. The registered Database name, or path to the Database file to set as a source.
;                  $iContentType        - [optional] an integer value (0-2). Default is Null. The type of Data to use from the Data Source. See Constants $LOW_FORM_CONTENT_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sContent            - [optional] a string value. Default is Null. The Content to be used for the form, either a Table or Query name, or an SQL statement.
;                  $bAnalyzeSQL         - [optional] a boolean value. Default is Null. If True, SQL commands will be analyzed by LibreOffice.
;                  $sFilter             - [optional] a string value. Default is Null. The SQL filter command.
;                  $sSort               - [optional] a string value. Default is Null. The SQL Sort command.
;                  $aLinkMaster         - [optional] an array of unknowns. Default is Null. An array of Master Field names. See remarks.
;                  $aLinkSlave          - [optional] an array of unknowns. Default is Null. An array of Slave Field names. See remarks.
;                  $bAdditions          - [optional] a boolean value. Default is Null. If True, additions are allowed.
;                  $bModifications      - [optional] a boolean value. Default is Null. If True, Modifications are allowed.
;                  $bDeletions          - [optional] a boolean value. Default is Null. If True, Deletions are allowed.
;                  $bAddOnly            - [optional] a boolean value. Default is Null. If True, Data can only be added.
;                  $iNavBar             - [optional] an integer value (0-2). Default is Null. The Navigation Bar mode. See Constants $LOW_FORM_NAV_BAR_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iCycle              - [optional] an integer value (0-2). Default is Null. What happens when you press the Tab key at the end of a record. See remarks. See Constants $LOW_FORM_CYCLE_MODE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oForm not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oForm not a Form Object.
;                  @Error 1 @Extended 3 Return 0 = $sSource not a String.
;                  @Error 1 @Extended 4 Return 0 = Source called in $sSource not a registered Database, and file does not exist.
;                  @Error 1 @Extended 5 Return 0 = $iContentType not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CONTENT_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $sContent not a String.
;                  @Error 1 @Extended 7 Return 0 = $bAnalyzeSQL not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $sFilter not a String.
;                  @Error 1 @Extended 9 Return 0 = $sSort not a String.
;                  @Error 1 @Extended 10 Return 0 = $aLinkMaster not an Array.
;                  @Error 1 @Extended 11 Return 0 = $aLinkSlave not an Array.
;                  @Error 1 @Extended 12 Return 0 = $bAdditions not a Boolean.
;                  @Error 1 @Extended 13 Return 0 = $bModifications not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $bDeletions not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $bAddOnly not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $iNavBar not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_NAV_BAR_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 17 Return 0 = $iCycle not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_CYCLE_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Data Source Name.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sSource
;                  |                               2 = Error setting $iContentType
;                  |                               4 = Error setting $sContent
;                  |                               8 = Error setting $bAnalyzeSQL
;                  |                               16 = Error setting $sFilter
;                  |                               32 = Error setting $sSort
;                  |                               64 = Error setting $aLinkMaster
;                  |                               128 = Error setting $aLinkSlave
;                  |                               256 = Error setting $bAdditions
;                  |                               512 = Error setting $bModifications
;                  |                               1024 = Error setting $bDeletions
;                  |                               2048 = Error setting $bAddOnly
;                  |                               4096 = Error setting $iNavBar
;                  |                               8192 = Error setting $iCycle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 14 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The array called for either $aLinkMaster or $aLinkSlave, should be a single dimension array, with one Field name per array element.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Error checking for $aLinkMaster or $aLinkSlave does not check content, only array size.
;                  There currently is no ability to set $iCycle to Default, but when it is already set to Default, the return value will be an empty string.
; Related .......: _LOWriter_FormPropertiesGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormPropertiesData(ByRef $oForm, $sSource = Null, $iContentType = Null, $sContent = Null, $bAnalyzeSQL = Null, $sFilter = Null, $sSort = Null, $aLinkMaster = Null, $aLinkSlave = Null, $bAdditions = Null, $bModifications = Null, $bDeletions = Null, $bAddOnly = Null, $iNavBar = Null, $iCycle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext
	Local $avForm[14]
	Local $iError = 0
	Local $sSourceName

	If Not IsObj($oForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oForm.supportsService("com.sun.star.form.component.Form") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sSource, $iContentType, $sContent, $bAnalyzeSQL, $sFilter, $sSort, $aLinkMaster, $aLinkSlave, $bAdditions, $bModifications, $bDeletions, $bAddOnly, $iNavBar, $iCycle) Then
		$sSourceName = $oForm.DataSourceName()
		If Not IsString($sSourceName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If StringInStr($sSourceName, "/") Then $sSourceName = _LO_PathConvert($sSourceName, $LO_PATHCONV_PCPATH_RETURN)

		__LO_ArrayFill($avForm, $sSourceName, $oForm.CommandType(), $oForm.Command(), $oForm.EscapeProcessing(), $oForm.Filter(), $oForm.Order(), _
				$oForm.MasterFields(), $oForm.DetailFields(), $oForm.AllowInserts(), $oForm.AllowUpdates(), $oForm.AllowDeletes(), $oForm.IgnoreResult(), $oForm.NavigationBarMode(), $oForm.Cycle())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avForm)
	EndIf

	If ($sSource <> Null) Then
		If Not IsString($sSource) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
		If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
		If Not ($oDBaseContext.hasByName($sSource)) And Not FileExists(_LO_PathConvert($sSource, $LO_PATHCONV_PCPATH_RETURN)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If StringInStr($sSource, "\") Then $sSource = _LO_PathConvert($sSource, $LO_PATHCONV_OFFICE_RETURN)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oForm.DataSourceName = $sSource
		$iError = ($oForm.DataSourceName() = $sSource) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iContentType <> Null) Then
		If Not __LO_IntIsBetween($iContentType, $LOW_FORM_CONTENT_TYPE_TABLE, $LOW_FORM_CONTENT_TYPE_SQL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oForm.CommandType = $iContentType
		$iError = ($oForm.CommandType() = $iContentType) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oForm.Command = $sContent
		$iError = ($oForm.Command() = $sContent) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bAnalyzeSQL <> Null) Then
		If Not IsBool($bAnalyzeSQL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oForm.EscapeProcessing = $bAnalyzeSQL
		$iError = ($oForm.EscapeProcessing() = $bAnalyzeSQL) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sFilter <> Null) Then
		If Not IsString($sFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oForm.Filter = $sFilter
		$iError = ($oForm.Filter() = $sFilter) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($sSort <> Null) Then
		If Not IsString($sSort) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oForm.Order = $sSort
		$iError = ($oForm.Order() = $sSort) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($aLinkMaster <> Null) Then
		If Not IsArray($aLinkMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oForm.MasterFields = $aLinkMaster
		$iError = (UBound($oForm.MasterFields()) = UBound($aLinkMaster)) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($aLinkSlave <> Null) Then
		If Not IsArray($aLinkSlave) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oForm.DetailFields = $aLinkSlave
		$iError = (UBound($oForm.DetailFields()) = UBound($aLinkSlave)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($bAdditions <> Null) Then
		If Not IsBool($bAdditions) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oForm.AllowInserts = $bAdditions
		$iError = ($oForm.AllowInserts() = $bAdditions) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($bModifications <> Null) Then
		If Not IsBool($bModifications) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$oForm.AllowUpdates = $bModifications
		$iError = ($oForm.AllowUpdates() = $bModifications) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($bDeletions <> Null) Then
		If Not IsBool($bDeletions) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$oForm.AllowDeletes = $bDeletions
		$iError = ($oForm.AllowDeletes() = $bDeletions) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($bAddOnly <> Null) Then
		If Not IsBool($bAddOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$oForm.IgnoreResult = $bAddOnly
		$iError = ($oForm.IgnoreResult() = $bAddOnly) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	If ($iNavBar <> Null) Then
		If Not __LO_IntIsBetween($iNavBar, $LOW_FORM_NAV_BAR_MODE_NO, $LOW_FORM_NAV_BAR_MODE_PARENT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$oForm.NavigationBarMode = $iNavBar
		$iError = ($oForm.NavigationBarMode() = $iNavBar) ? ($iError) : (BitOR($iError, 4096))
	EndIf

	If ($iCycle <> Null) Then
		If Not __LO_IntIsBetween($iCycle, $LOW_FORM_CYCLE_MODE_ALL_RECORDS, $LOW_FORM_CYCLE_MODE_CURRENT_PAGE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

		$oForm.Cycle = $iCycle
		$iError = ($oForm.Cycle() = $iCycle) ? ($iError) : (BitOR($iError, 8192))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormPropertiesData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormPropertiesGeneral
; Description ...: Set or Retrieve Form general properties.
; Syntax ........: _LOWriter_FormPropertiesGeneral(ByRef $oForm[, $sName = Null[, $sURL = Null[, $sFrame = Null[, $iEncoding = Null[, $iSubType = Null]]]]])
; Parameters ....: $oForm               - [in/out] an object. A Form object returned by a previous _LOWriter_FormAdd, _LOWriter_FormGetObjByIndex or _LOWriter_FormsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The name of the Form.
;                  $sURL                - [optional] a string value. Default is Null. The URL or Document path to open.
;                  $sFrame              - [optional] a string value. Default is Null. The frame to open the URL in. See Constants, $LOW_FRAME_TARGET_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iEncoding           - [optional] an integer value (0-2). Default is Null. The type of encoding for the data transfer. See Constants $LOW_FORM_SUBMIT_ENCODING_* as defined in LibreOfficeWriter_Constants.au3
;                  $iSubType            - [optional] an integer value (0-1). Default is Null. The method to submit the completed form. See Constants $LOW_FORM_SUBMIT_METHOD_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oForm not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oForm not a Form Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sURL not a String.
;                  @Error 1 @Extended 5 Return 0 = $sFrame not a String.
;                  @Error 1 @Extended 6 Return 0 = $sFrame not called with correct constant. See Constants, $LOW_FRAME_TARGET_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $iEncoding not an Integer, less than 0 or greater than 2. See Constants $LOW_FORM_SUBMIT_ENCODING_* as defined in LibreOfficeWriter_Constants.au3
;                  @Error 1 @Extended 8 Return 0 = $iSubType not an Integer, less than 0 or greater than 1. See Constants $LOW_FORM_SUBMIT_METHOD_* as defined in LibreOfficeWriter_Constants.au3
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sURL
;                  |                               4 = Error setting $sFrame
;                  |                               8 = Error setting $iEncoding
;                  |                               16 = Error setting $iSubType
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FormPropertiesData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormPropertiesGeneral(ByRef $oForm, $sName = Null, $sURL = Null, $sFrame = Null, $iEncoding = Null, $iSubType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avForm[5]
	Local $iError = 0

	If Not IsObj($oForm) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oForm.supportsService("com.sun.star.form.component.Form") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sName, $sURL, $sFrame, $iEncoding, $iSubType) Then
		__LO_ArrayFill($avForm, $oForm.Name(), $oForm.URL(), $oForm.TargetFrame(), $oForm.SubmitEncoding(), $oForm.SubmitMethod())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avForm)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oForm.Name = $sName
		$iError = ($oForm.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sURL <> Null) Then
		If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oForm.URL = $sURL
		$iError = ($oForm.URL() = $sURL) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sFrame <> Null) Then
		If Not IsString($sFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If ($sFrame <> $LOW_FRAME_TARGET_TOP) And _
				($sFrame <> $LOW_FRAME_TARGET_PARENT) And _
				($sFrame <> $LOW_FRAME_TARGET_BLANK) And _
				($sFrame <> $LOW_FRAME_TARGET_SELF) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oForm.TargetFrame = $sFrame
		$iError = ($oForm.TargetFrame() = $sFrame) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iEncoding <> Null) Then
		If Not __LO_IntIsBetween($iEncoding, $LOW_FORM_SUBMIT_ENCODING_URL, $LOW_FORM_SUBMIT_ENCODING_TEXT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oForm.SubmitEncoding = $iEncoding
		$iError = ($oForm.SubmitEncoding() = $iEncoding) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iSubType <> Null) Then
		If Not __LO_IntIsBetween($iSubType, $LOW_FORM_SUBMIT_METHOD_GET, $LOW_FORM_SUBMIT_METHOD_POST) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oForm.SubmitMethod = $iSubType
		$iError = ($oForm.SubmitMethod() = $iSubType) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FormPropertiesGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormsGetCount
; Description ...: Retrieve a count of top-level forms contained in a document or sub-forms of a form.
; Syntax ........: _LOWriter_FormsGetCount(ByRef $oObj)
; Parameters ....: $oObj                - [in/out] an object. Either a Document Object or a Form object. See Remarks. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function, or a Form Object returned from a previous _LOWriter_FormsGetList, or _LOWriter_FormAdd function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = Called Object in $oObj, not a Document Object, and not a Form Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a count of forms.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning a count of Forms contained in the Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $oObj is called with a Document object, a count of top level forms will be returned. If a Form object is called, a count of all sub-forms for the form wil be returned.
; Related .......: _LOWriter_FormsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormsGetCount(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If $oObj.supportsService("com.sun.star.form.component.Form") Then
		For $i = 0 To $oObj.Count() - 1
			If $oObj.getByIndex($i).supportsService("com.sun.star.form.component.Form") Then $iCount += 1
		Next

	ElseIf $oObj.supportsService("com.sun.star.text.TextDocument") Then
		$iCount = $oObj.DrawPage.Forms.Count()
		If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; wrong type of input item.
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOWriter_FormsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormsGetList
; Description ...: Retrieve an array of Form Objects.
; Syntax ........: _LOWriter_FormsGetList(ByRef $oObj)
; Parameters ....: $oObj                - [in/out] an object. Either a Document Object or a Form object. See Remarks. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function, or a Form Object returned from a previous _LOWriter_FormsGetList, or _LOWriter_FormAdd function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = Called Object in $oObj, not a Document Object, and not a Form Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve form Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an array of Form Objects. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $oObj is called with a Document object, an array of top level forms will be returned. If a Form object is called, all sub-forms for the form will be returned.
; Related .......: _LOWriter_FormAdd
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormsGetList(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local $aoForms[0]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If $oObj.supportsService("com.sun.star.form.component.Form") Then
		For $i = 0 To $oObj.Count() - 1
			If $oObj.getByIndex($i).supportsService("com.sun.star.form.component.Form") Then
				ReDim $aoForms[$iCount + 1]
				$aoForms[$iCount] = $oObj.getByIndex($i)
				$iCount += 1
			EndIf
		Next

	ElseIf $oObj.supportsService("com.sun.star.text.TextDocument") Then
		ReDim $aoForms[$oObj.DrawPage.Forms.Count()]

		For $i = 0 To $oObj.DrawPage.Forms.Count() - 1
			$aoForms[$i] = $oObj.DrawPage.Forms.getByIndex($i)
			If Not IsObj($aoForms[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		Next

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; wrong type of input item.
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoForms), $aoForms)
EndFunc   ;==>_LOWriter_FormsGetList
