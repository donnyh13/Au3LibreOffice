#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Internal.au3"

; Other includes for Calc

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Removing, etc. L.O. Calc document Fields.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_FieldCurrentDisplayGet
; _LOCalc_FieldDateTimeInsert
; _LOCalc_FieldDelete
; _LOCalc_FieldFileNameInsert
; _LOCalc_FieldGetAnchor
; _LOCalc_FieldHyperlinkInsert
; _LOCalc_FieldHyperlinkModify
; _LOCalc_FieldPageCountInsert
; _LOCalc_FieldPageNumberInsert
; _LOCalc_FieldsGetList
; _LOCalc_FieldSheetNameInsert
; _LOCalc_FieldTitleInsert
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldCurrentDisplayGet
; Description ...: Retrieve the current data displayed by a field.
; Syntax ........: _LOCalc_FieldCurrentDisplayGet(ByRef $mField[, $bFieldName = False])
; Parameters ....: $mField              - [in/out] a map. A Map containing a Field Object as returned from a previous insert, or _LOCalc_FieldsGetList function.
;                  $bFieldName          - [optional] a boolean value. Default is False. If True, the Field command name is returned, else the current Field display is returned.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oField not a map.
;                  @Error 1 @Extended 2 Return 0 = $bFieldName not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Field's current display.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify alternate Field Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning current Field display content in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving the current display of Fields in the Header/Footer, generally returns three question marks, this is due to the way Headers/Footers are implemented for Calc.
;                  Retrieving the current Field display will return the display at the time of Field Object retrieval.
; Related .......: _LOCalc_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldCurrentDisplayGet(ByRef $mField, $bFieldName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sPresentation

	If Not IsMap($mField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bFieldName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sPresentation = ($mField.EnumFieldObj).getPresentation($bFieldName)
	If Not IsString($sPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sPresentation)
EndFunc   ;==>_LOCalc_FieldCurrentDisplayGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldDateTimeInsert
; Description ...: Insert a Date or Time Field.
; Syntax ........: _LOCalc_FieldDateTimeInsert(ByRef $oDoc, ByRef $oTextCursor[, $bIsDate = True[, $bOverwrite = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bIsDate             - [optional] a boolean value. Default is True. If True, the inserted Field will be a Date Field, if False, the Field will be a Time Field.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bIsDate not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a "com.sun.star.text.TextField.DateTime" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify and retrieve Field object after insertion.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Successfully inserted the field, returning a Map containing the Field's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you are inserting the field into a header or footer of the document, make sure you use a newly created Text Cursor, using a cursor that has previously inserted text, will cause this function to fail to identify the new Field's object. The Field will still be successfully inserted however.
;                  Inserting Fields into the Headers/Footers is a bit glitchy.
;                  The reason I use a Map to contain the Field's object is that Calc Fields are a little buggy currently, therefore I need two Objects in order to do certain functions with the Field, such as deleteing, or retrieving the Field's display. It is easier and more accurate to identify and retrieve the Objects now, rather than later.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldDateTimeInsert(ByRef $oDoc, ByRef $oTextCursor, $bIsDate = True, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bIsDate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.DateTime")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	With $oTextField
		.IsFixed = False
		.IsDate = $bIsDate
		.NumberFormat = 2
	EndWith

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	$oTextFieldReturn = __LOCalc_FieldGetObj($oTextCursor, $LOC_FIELD_TYPE_DATE_TIME)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOCalc_FieldDateTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldDelete
; Description ...: Delete a Field from a Document.
; Syntax ........: _LOCalc_FieldDelete(ByRef $mField)
; Parameters ....: $mField              - [in/out] a map. A Map containing a Field Object as returned from a previous insert, or _LOCalc_FieldsGetList function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $mField not a map.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a new Text Cursor.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully deleted the field.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To Delete a field in a Header/Footer, retrieve a new Object for the Header/Footer, retrieve an array of fields from the Header/Footer Object, delete the Field using this function, then re-insert the Header/Footer object into the Page Style.
; Related .......: _LOCalc_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldDelete(ByRef $mField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCursor

	If Not IsMap($mField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($mField["FieldObj"].Anchor.Text.SupportsService("com.sun.star.sheet.SheetCell")) Then
		$mField["FieldObj"].Anchor.Text.removeTextContent($mField["FieldObj"])

	Else
		$oCursor = $mField["FieldObj"].Anchor.Text.createTextCursorByRange($mField["EnumFieldObj"].Anchor())
		If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oCursor.Text.insertString($oCursor, "", True)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_FieldDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldFileNameInsert
; Description ...: Insert a File Name field.
; Syntax ........: _LOCalc_FieldFileNameInsert(ByRef $oDoc, ByRef $oTextCursor[, $bPath = False[, $bOverwrite = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bPath               - [optional] a boolean value. Default is False. If True, the File name will be prefixed by the File Path. If False, the File name and extension will be displayed.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bPath not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a "com.sun.star.text.TextField.FileName" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify and retrieve Field object after insertion.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Successfully inserted the field, returning a map containing the Field's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you are inserting the field into a header or footer of the document, make sure you use a newly created Text Cursor, using a cursor that has previously inserted text, will cause this function to fail to identify the new Field's object. The Field will still be successfully inserted however.
;                  Inserting Fields into the Headers/Footers is a bit glitchy.
;                  The reason I use a Map to contain the Field's object is that Calc Fields are a little buggy currently, therefore I need two Objects in order to do certain functions with the Field, such as deleteing, or retrieving the Field's display. It is easier and more accurate to identify and retrieve the Objects now, rather than later.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldFileNameInsert(ByRef $oDoc, ByRef $oTextCursor, $bPath = False, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__LOCCONST_FilenameDisplayFormat_FULL = 0, $__LOCCONST_FilenameDisplayFormat_NAME_AND_EXT = 3
	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bPath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.FileName")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTextField.FileFormat = ($bPath) ? ($__LOCCONST_FilenameDisplayFormat_FULL) : ($__LOCCONST_FilenameDisplayFormat_NAME_AND_EXT)

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	$oTextFieldReturn = __LOCalc_FieldGetObj($oTextCursor, $LOC_FIELD_TYPE_FILE_NAME)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOCalc_FieldFileNameInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldGetAnchor
; Description ...: Retrieve the Anchor Cursor Object for a Field.
; Syntax ........: _LOCalc_FieldGetAnchor(ByRef $mField)
; Parameters ....: $mField              - [in/out] a map. A Map containing a Field Object as returned from a previous insert, or _LOCalc_FieldsGetList function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $mField not a map.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to retrieve Field anchor Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Field Anchor Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldGetAnchor(ByRef $mField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFieldAnchor

	If Not IsMap($mField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oFieldAnchor = ($mField.FieldObj).Anchor.Text.createTextCursorByRange(($mField.EnumFieldObj).Anchor())
	If Not IsObj($oFieldAnchor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFieldAnchor)
EndFunc   ;==>_LOCalc_FieldGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldHyperlinkInsert
; Description ...: Insert a Hyperlink into a Calc Cell.
; Syntax ........: _LOCalc_FieldHyperlinkInsert(ByRef $oDoc, ByRef $oTextCursor, $sURL[, $sText = ""[, $sTargetFrame = ""[, $bOverwrite = False]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $sURL                - a string value. The URL/Hyperlink Address.
;                  $sText               - [optional] a string value. Default is "". The Text to display instead of the URL. "" means the URL itself is displayed.
;                  $sTargetFrame        - [optional] a string value. Default is "". Enter the name of the frame that you want the linked file to open in. Leave blank to skip.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sURL not a String.
;                  @Error 1 @Extended 4 Return 0 = $sText not a String.
;                  @Error 1 @Extended 5 Return 0 = $sTargetFrame not a String.
;                  @Error 1 @Extended 6 Return 0 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a "com.sun.star.text.TextField.URL" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify and retrieve Field object after insertion.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Successfully inserted the field, returning a map containing the Field's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The reason I use a Map to contain the Field's object is that Calc Fields are a little buggy currently, therefore I need two Objects in order to do certain functions with the Field, such as deleteing, or retrieving the Field's display. It is easier and more accurate to identify and retrieve the Objects now, rather than later.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldHyperlinkInsert(ByRef $oDoc, ByRef $oTextCursor, $sURL, $sText = "", $sTargetFrame = "", $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
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

	$oTextFieldReturn = __LOCalc_FieldGetObj($oTextCursor, $LOC_FIELD_TYPE_URL)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOCalc_FieldHyperlinkInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldHyperlinkModify
; Description ...: Set or Retrieve the settings of a Hyperlink/URL field.
; Syntax ........: _LOCalc_FieldHyperlinkModify(ByRef $mHyperlinkField[, $sURL = Null[, $sText = Null[, $sTargetFrame = Null]]])
; Parameters ....: $mHyperlinkField     - [in/out] a map. A Hyperlink/URL Map containing the Field Field Object returned by a previous _LOCalc_FieldHyperlinkInsert or _LOCalc_FieldsGetList function.
;                  $sURL                - [optional] a string value. Default is Null. The URL/Hyperlink Address.
;                  $sText               - [optional] a string value. Default is Null. The Text to display instead of the URL. "" means the URL itself is displayed.
;                  $sTargetFrame        - [optional] a string value. Default is Null. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $mHyperlinkField not a map.
;                  @Error 1 @Extended 2 Return 0 = $sURL not a String.
;                  @Error 1 @Extended 3 Return 0 = $sText not a String.
;                  @Error 1 @Extended 4 Return 0 = $sTargetFrame not a String.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sURL
;                  |                               2 = Error setting $sText
;                  |                               4 = Error setting $sTargetFrame
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldHyperlinkModify(ByRef $mHyperlinkField, $sURL = Null, $sText = Null, $sTargetFrame = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asField[3]

	If Not IsMap($mHyperlinkField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sURL, $sText, $sTargetFrame) Then
		__LO_ArrayFill($asField, $mHyperlinkField["FieldObj"].URL(), $mHyperlinkField["FieldObj"].Representation(), $mHyperlinkField["FieldObj"].TargetFrame())

		Return SetError($__LO_STATUS_SUCCESS, 1, $asField)
	EndIf

	If ($sURL <> Null) Then
		If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$mHyperlinkField["FieldObj"].URL = $sURL
		$iError = ($mHyperlinkField["FieldObj"].URL() = $sURL) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$mHyperlinkField["FieldObj"].Representation = $sText
		$iError = ($mHyperlinkField["FieldObj"].Representation() = $sText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sTargetFrame <> Null) Then
		If Not IsString($sTargetFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$mHyperlinkField["FieldObj"].TargetFrame = $sTargetFrame
		$iError = ($mHyperlinkField["FieldObj"].TargetFrame() = $sTargetFrame) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_FieldHyperlinkModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldPageCountInsert
; Description ...: Insert a total Page Count Field.
; Syntax ........: _LOCalc_FieldPageCountInsert(ByRef $oDoc, ByRef $oTextCursor[, $bOverwrite = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a "com.sun.star.text.TextField.PageCount" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify and retrieve Field object after insertion.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Successfully inserted the field, returning a map containing the Field's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you are inserting the field into a header or footer of the document, make sure you use a newly created Text Cursor, using a cursor that has previously inserted text, will cause this function to fail to identify the new Field's object. The Field will still be successfully inserted however.
;                  Inserting Fields into the Headers/Footers is a bit glitchy.
;                  The reason I use a Map to contain the Field's object is that Calc Fields are a little buggy currently, therefore I need two Objects in order to do certain functions with the Field, such as deleteing, or retrieving the Field's display. It is easier and more accurate to identify and retrieve the Objects now, rather than later.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldPageCountInsert(ByRef $oDoc, ByRef $oTextCursor, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.PageCount")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	$oTextFieldReturn = __LOCalc_FieldGetObj($oTextCursor, $LOC_FIELD_TYPE_PAGE_COUNT)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOCalc_FieldPageCountInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldPageNumberInsert
; Description ...: Insert a Page Number Field.
; Syntax ........: _LOCalc_FieldPageNumberInsert(ByRef $oDoc, ByRef $oTextCursor[, $bOverwrite = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a "com.sun.star.text.TextField.PageNumber" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify and retrieve Field object after insertion.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Successfully inserted the field, returning a map containing the Field's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you are inserting the field into a header or footer of the document, make sure you use a newly created Text Cursor, using a cursor that has previously inserted text, will cause this function to fail to identify the new Field's object. The Field will still be successfully inserted however.
;                  Inserting Fields into the Headers/Footers is a bit glitchy.
;                  The reason I use a Map to contain the Field's object is that Calc Fields are a little buggy currently, therefore I need two Objects in order to do certain functions with the Field, such as deleteing, or retrieving the Field's display. It is easier and more accurate to identify and retrieve the Objects now, rather than later.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldPageNumberInsert(ByRef $oDoc, ByRef $oTextCursor, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.PageNumber")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	$oTextFieldReturn = __LOCalc_FieldGetObj($oTextCursor, $LOC_FIELD_TYPE_PAGE_NUM)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOCalc_FieldPageNumberInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldsGetList
; Description ...: Retrieve an Array of maps containing Field Objects present in a Cell or Header/Footer.
; Syntax ........: _LOCalc_FieldsGetList(ByRef $oTextCursor[, $iType = $LOC_FIELD_TYPE_ALL[, $bFieldTypeNum = True]])
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $iType               - [optional] an integer value (1-255). Default is $LOC_FIELD_TYPE_ALL. The type of Field to search for. See Constants, $LOC_FIELD_TYPE_* as defined in LibreOfficeCalc_Constants.au3. Can be BitOr'd together.
;                  $bFieldTypeNum       - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type Constant Integer for that particular Field, to assist in identifying the Field type. See Constants, $LOC_FIELD_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 1, or greater than 255. (The total of all Constants added together.) See Constants, $LOC_FIELD_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $bFieldTypeNum not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create enumeration of paragraphs in Cell.
;                  @Error 2 @Extended 2 Return 0 = Failed to create enumeration of Text Portions in Paragraph.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify requested Field Types.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Text Fields Object/
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve total count of Fields.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Text Field Object.
;                  @Error 3 @Extended 5 Return 0 = More fields found than total count of Fields. Try creating a new cursor, and trying again.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve secondary Text Field Object.
;                  @Error 3 @Extended 7 Return 0 = Number of Fields found not equal to number of expected Fields. Try creating a new cursor, and trying again.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of maps containing Text Field Objects with @Extended set to number of results. See Remarks for Array sizing.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array can vary in the number of columns, if $bFieldTypeNum is set to False, the Array will be a single column. If $bFieldTypeNum is set to True, a column will be added to the array. First column will always be the amp containing the Field's Object.
;                  Setting $bFieldTypeNum to True will add a Field type Number column, matching the constants, $LOC_FIELD_TYPE_* as defined in LibreOfficeCalc_Constants.au3 for the found Field.
;                  This function may fail to identify Fields if text has been inserted recently using the same Cursor.
;                  The reason I use a Map to contain the Field's object is that Calc Fields are a little buggy currently, therefore I need two Objects in order to do certain functions with the Field, such as deleteing, or retrieving the Field's display. It is easier and more accurate to identify and retrieve the Objects now, rather than later.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldsGetList(ByRef $oTextCursor, $iType = $LOC_FIELD_TYPE_ALL, $bFieldTypeNum = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avFieldTypes[0][0]
	Local $oParEnum, $oPar, $oTextEnum, $oTextPortion, $oTextField, $oInternalCursor = $oTextCursor, $oFields, $oField
	Local $iCount = 0, $iTotalFound = 0, $iTotalFields = 0
	Local $mFieldObj[]
	Local $avTextFields[1]

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iType, $LOC_FIELD_TYPE_ALL, 255) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	; 255 is all possible Consts added together
	If Not IsBool($bFieldTypeNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	; When a Text Cursor has been used to insert Strings previous to inserting or looking for a Field, the fields sometimes are not able to be identified.
	; The workaround I figured out was to create the Text Cursor again before enumerating the fields. I only create the text cursor again if the Text Cursor is in a Cell, not a header.
	If ($oTextCursor.Text.SupportsService("com.sun.star.sheet.SheetCell")) Then
		$oInternalCursor = $oTextCursor.Text.Spreadsheet.getCellByPosition($oTextCursor.Text.RangeAddress.StartColumn(), $oTextCursor.Text.RangeAddress.StartRow()).Text.createTextCursor()
	EndIf

	$avFieldTypes = __LOCalc_FieldTypeServices($iType)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $bFieldTypeNum Then ReDim $avTextFields[1][2]

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
					If $oTextField.supportsService($avFieldTypes[$i][1]) Then
						$oField = $oFields.getByIndex($iTotalFound)
						If Not IsObj($oField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

						If $bFieldTypeNum Then
							$mFieldObj.EnumFieldObj = $oTextField
							$mFieldObj.FieldObj = $oField
							$avTextFields[$iCount][0] = $mFieldObj
							$avTextFields[$iCount][1] = $avFieldTypes[$i][0]
							$iCount += 1
							If ($iCount = UBound($avTextFields)) Then ReDim $avTextFields[$iCount * 2][2]

						Else
							$mFieldObj.EnumFieldObj = $oTextField
							$mFieldObj.FieldObj = $oField
							$avTextFields[$iCount] = $mFieldObj
							$iCount += 1
							If ($iCount = UBound($avTextFields)) Then ReDim $avTextFields[$iCount * 2]
						EndIf

						ExitLoop
					EndIf
					Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
				Next

				$iTotalFound += 1
			EndIf
		WEnd
	WEnd

	If $bFieldTypeNum Then
		ReDim $avTextFields[$iCount][2]

	Else
		ReDim $avTextFields[$iCount]
	EndIf

	If $iTotalFields <> $iTotalFound Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $avTextFields)
EndFunc   ;==>_LOCalc_FieldsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldSheetNameInsert
; Description ...: Insert a Sheet Name Field at a Text Cursor location.
; Syntax ........: _LOCalc_FieldSheetNameInsert(ByRef $oDoc, ByRef $oTextCursor[, $bOverwrite = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a "com.sun.star.text.TextField.SheetName" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify and retrieve Field object after insertion.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Successfully inserted the field, returning a map containing the Field's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you are inserting the field into a header or footer of the document, make sure you use a newly created Text Cursor, using a cursor that has previously inserted text, will cause this function to fail to identify the new Field's object. The Field will still be successfully inserted however.
;                  Inserting Fields into the Headers/Footers is a bit glitchy.
;                  The reason I use a Map to contain the Field's object is that Calc Fields are a little buggy currently, therefore I need two Objects in order to do certain functions with the Field, such as deleteing, or retrieving the Field's display. It is easier and more accurate to identify and retrieve the Objects now, rather than later.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldSheetNameInsert(ByRef $oDoc, ByRef $oTextCursor, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.SheetName")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	$oTextFieldReturn = __LOCalc_FieldGetObj($oTextCursor, $LOC_FIELD_TYPE_SHEET_NAME)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOCalc_FieldSheetNameInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FieldTitleInsert
; Description ...: Insert a Document Title field.
; Syntax ........: _LOCalc_FieldTitleInsert(ByRef $oDoc, ByRef $oTextCursor[, $bOverwrite = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a "com.sun.star.text.TextField.docinfo.Title" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify and retrieve Field object after insertion.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Successfully inserted the field, returning a map containing the Field's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you are inserting the field into a header or footer of the document, make sure you use a newly created Text Cursor, using a cursor that has previously inserted text, will cause this function to fail to identify the new Field's object. The Field will still be successfully inserted however.
;                  Inserting Fields into the Headers/Footers is a bit glitchy.
;                  The reason I use a Map to contain the Field's object is that Calc Fields are a little buggy currently, therefore I need two Objects in order to do certain functions with the Field, such as deleteing, or retrieving the Field's display. It is easier and more accurate to identify and retrieve the Objects now, rather than later.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FieldTitleInsert(ByRef $oDoc, ByRef $oTextCursor, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.docinfo.Title")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	$oTextFieldReturn = __LOCalc_FieldGetObj($oTextCursor, $LOC_FIELD_TYPE_DOC_TITLE)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOCalc_FieldTitleInsert
