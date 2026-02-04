#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Base
#include "LibreOfficeBase_Internal.au3"

; Other includes for Base

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Adding, Deleting, and modifying, etc. L.O. Base Reports.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; Notes .........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_ReportClose
; _LOBase_ReportConDelete
; _LOBase_ReportConFormattedFieldData
; _LOBase_ReportConFormattedFieldGeneral
; _LOBase_ReportConImageConData
; _LOBase_ReportConImageConGeneral
; _LOBase_ReportConInsert
; _LOBase_ReportConLabelGeneral
; _LOBase_ReportConLineGeneral
; _LOBase_ReportConnect
; _LOBase_ReportConPosition
; _LOBase_ReportConsGetList
; _LOBase_ReportConSize
; _LOBase_ReportCopy
; _LOBase_ReportCreate
; _LOBase_ReportData
; _LOBase_ReportDelete
; _LOBase_ReportDetail
; _LOBase_ReportDocVisible
; _LOBase_ReportExists
; _LOBase_ReportFolderCopy
; _LOBase_ReportFolderCreate
; _LOBase_ReportFolderDelete
; _LOBase_ReportFolderExists
; _LOBase_ReportFolderRename
; _LOBase_ReportFoldersGetCount
; _LOBase_ReportFoldersGetNames
; _LOBase_ReportFooter
; _LOBase_ReportGeneral
; _LOBase_ReportGroupAdd
; _LOBase_ReportGroupDeleteByIndex
; _LOBase_ReportGroupDeleteByObj
; _LOBase_ReportGroupFooter
; _LOBase_ReportGroupGetByIndex
; _LOBase_ReportGroupHeader
; _LOBase_ReportGroupPosition
; _LOBase_ReportGroupsGetCount
; _LOBase_ReportGroupSort
; _LOBase_ReportHeader
; _LOBase_ReportIsModified
; _LOBase_ReportOpen
; _LOBase_ReportPageFooter
; _LOBase_ReportPageHeader
; _LOBase_ReportRename
; _LOBase_ReportSave
; _LOBase_ReportSectionGetObj
; _LOBase_ReportsGetCount
; _LOBase_ReportsGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportClose
; Description ...: Close an opened Report Document.
; Syntax ........: _LOBase_ReportClose(ByRef $oReportDoc[, $bForceClose = False])
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $bForceClose         - [optional] a boolean value. Default is False. If True, the Report document will be closed regardless if there are unsaved changes. See remarks.
; Return values .: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bForceClose not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Document has been modified and not saved, and $bForceClose is False.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Report Document's properties.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify Report in Parent Document.
;                  @Error 3 @Extended 5 Return 0 = Document called in $oReportDoc not a Report Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning a Boolean value of whether the Report Document was successfully closed (True), or not.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If there are unsaved changes in the document when close is called, and $bForceClose is True, they will be lost.
; Related .......: _LOBase_ReportOpen, _LOBase_ReportConnect, _LOBase_ReportDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportClose(ByRef $oReportDoc, $bForceClose = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn
	Local $oReport, $oSource
	Local $tPropertiesPair

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bForceClose) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oReportDoc.isModified() And Not $bForceClose Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oReportDoc.supportsService("com.sun.star.text.TextDocument") Then ; Report Doc is in viewing/Read-Only mode.
		$oReportDoc.close(True)
		$bReturn = True

	ElseIf $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then  ; Report is in Design mode.
		$oSource = $oReportDoc.Parent.ReportDocuments()
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$tPropertiesPair = $oSource.Parent.CurrentController.identifySubComponent($oReportDoc)
		If Not IsObj($tPropertiesPair) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oReport = $oSource.getByHierarchicalName($tPropertiesPair.Second())
		If Not IsObj($oReport) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oReportDoc.isModified() Then $oReportDoc.Modified = False ; Set modified to false, so the user wont be prompted.

		$bReturn = $oReport.Close()

	Else ; Error, unknown document?

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_ReportClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConDelete
; Description ...: Delete a Report Control.
; Syntax ........: _LOBase_ReportConDelete(ByRef $oControl)
; Parameters ....: $oControl            - [in/out] an object. A Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Control's parent.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve parent document.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Control was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_ReportConInsert, _LOBase_ReportConsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConDelete(ByRef $oControl)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oParent

	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oParent = $oControl.Parent() ; Identify the parent document.
	If Not IsObj($oParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oParent.supportsService("com.sun.star.report.Section") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oParent.remove($oControl)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportConDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConFormattedFieldData
; Description ...: Set or Retrieve Formatted Field Data Properties.
; Syntax ........: _LOBase_ReportConFormattedFieldData(ByRef $oFormatField[, $sDataField = Null])
; Parameters ....: $oFormatField        - [in/out] an object. A Formatted Field Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The DataField value, see Remarks.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormatField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oFormatField not a Formatted Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning current setting as a String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
;                  DataField is a String that determines the content to be shown. The entry format is either of the following:
;                  - To display the value of a column, you would call $sDataField with field:[??] where "??" represents the column's name. e.g. field:[Unique_ID].
;                  - To display the result of a function, you would call $sDataField with rpt:[??] where "??" represents the function name. e.g. rpt:[MaximumUnique_IDReport].
;                  - According to the "XReportControlModel" documentation, the following expression is also acceptable: rpt:24+24-47.
; Related .......: _LOBase_ReportConFormattedFieldGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConFormattedFieldData(ByRef $oFormatField, $sDataField = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oFormatField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOBase_ReportConIdentify($oFormatField) <> $LOB_REP_CON_TYPE_FORMATTED_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField) Then

		Return SetError($__LO_STATUS_SUCCESS, 1, $oFormatField.DataField())
	EndIf

	If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oFormatField.DataField = $sDataField
	$iError = ($oFormatField.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportConFormattedFieldData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConFormattedFieldGeneral
; Description ...: Set or Retrieve general Formatted Field properties.
; Syntax ........: _LOBase_ReportConFormattedFieldGeneral(ByRef $oFormatField[, $sName = Null[, $sCondPrint = Null[, $bPrintRep = Null[, $bPrintRepOnGroup = Null[, $iBackColor = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null[, $iFormat = Null]]]]]]]]])
; Parameters ....: $oFormatField        - [in/out] an object. A Formatted Field Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $sCondPrint          - [optional] a string value. Default is Null. The conditional print expression, prefixed by "rpt:".
;                  $bPrintRep           - [optional] a boolean value. Default is Null. If True, repeated values will be printed.
;                  $bPrintRepOnGroup    - [optional] a boolean value. Default is Null. If True, repeated values will be printed on group change.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF to set Background color to default / Background Transparent = True.
;                  $mFont               - [optional] a map. Default is Null. A Font descriptor Map returned by a previous _LOBase_FontDescCreate or _LOBase_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOB_TXT_ALIGN_HORI_* as defined in LibreOfficeBase_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOB_ALIGN_VERT_* as defined in LibreOfficeBase_Constants.au3.
;                  $iFormat             - [optional] an integer value. Default is Null. The Number Format Key to display the content in, retrieved from a previous _LOBase_FormatKeysGetList call, or created by _LOBase_FormatKeyCreate function.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oFormatField not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oFormatField not a Formatted Field Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sCondPrint not a String.
;                  @Error 1 @Extended 5 Return 0 = $bPrintRep not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bPrintRepOnGroup not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 8 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 9 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOB_TXT_ALIGN_HORI_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOB_ALIGN_VERT_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = $iFormat not an Integer.
;                  @Error 1 @Extended 12 Return 0 = Format key called in $iFormat not found in document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify Parent Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sCondPrint
;                  |                               4 = Error setting $bPrintRep
;                  |                               8 = Error setting $bPrintRepOnGroup
;                  |                               16 = Error setting $iBackColor
;                  |                               32 = Error setting $mFont
;                  |                               64 = Error setting $iAlign
;                  |                               128 = Error setting $iVertAlign
;                  |                               256 = Error setting $iFormat
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 10 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  I could not find a property to set the TextDirection or Visible settings.
;                  Background Transparent is set automatically based on the value set for Background color. Set Background color to $LO_COLOR_OFF to set Background Transparent to True.
; Related .......: _LOBase_FormatKeyCreate, _LOBase_FormatKeysGetList, _LOBase_ReportConFormattedFieldData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConFormattedFieldGeneral(ByRef $oFormatField, $sName = Null, $sCondPrint = Null, $bPrintRep = Null, $bPrintRepOnGroup = Null, $iBackColor = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oDoc
	Local $avControl[10]

	If Not IsObj($oFormatField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOBase_ReportConIdentify($oFormatField) <> $LOB_REP_CON_TYPE_FORMATTED_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sCondPrint, $bPrintRep, $bPrintRepOnGroup, $iBackColor, $mFont, $iAlign, $iVertAlign, $iFormat) Then
		__LO_ArrayFill($avControl, $oFormatField.Name(), $oFormatField.ConditionalPrintExpression(), $oFormatField.PrintRepeatedValues(), _
				$oFormatField.PrintWhenGroupChange(), _
				$oFormatField.ControlBackground(), __LOBase_ReportConSetGetFontDesc($oFormatField), $oFormatField.ParaAdjust(), _
				$oFormatField.VerticalAlign(), $oFormatField.FormatKey())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFormatField.Name = $sName
		$iError = ($oFormatField.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sCondPrint <> Null) Then
		If Not IsString($sCondPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oFormatField.ConditionalPrintExpression = $sCondPrint
		$iError = ($oFormatField.ConditionalPrintExpression() = $sCondPrint) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bPrintRep <> Null) Then
		If Not IsBool($bPrintRep) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFormatField.PrintRepeatedValues = $bPrintRep
		$iError = ($oFormatField.PrintRepeatedValues() = $bPrintRep) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bPrintRepOnGroup <> Null) Then
		If Not IsBool($bPrintRepOnGroup) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFormatField.PrintWhenGroupChange = $bPrintRepOnGroup
		$iError = ($oFormatField.PrintWhenGroupChange = $bPrintRepOnGroup) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFormatField.ControlBackground = $iBackColor
		$iError = ($oFormatField.ControlBackground() = $iBackColor) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		__LOBase_ReportConSetGetFontDesc($oFormatField, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOB_TXT_ALIGN_HORI_LEFT, $LOB_TXT_ALIGN_HORI_CENTER) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oFormatField.ParaAdjust = $iAlign
		$iError = ($oFormatField.ParaAdjust() = $iAlign) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOB_ALIGN_VERT_TOP, $LOB_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oFormatField.VerticalAlign = $iVertAlign
		$iError = ($oFormatField.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iFormat <> Null) Then
		If Not IsInt($iFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oDoc = $oFormatField.Parent() ; Identify the parent document.
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Do
			$oDoc = $oDoc.getParent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Until $oDoc.supportsService("com.sun.star.report.ReportDefinition")
		If Not _LOBase_FormatKeyExists($oDoc, $iFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oFormatField.FormatKey = $iFormat
		$iError = ($oFormatField.FormatKey() = $iFormat) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportConFormattedFieldGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConImageConData
; Description ...: Set or Retrieve Image Control Data Properties.
; Syntax ........: _LOBase_ReportConImageConData(ByRef $oImageControl[, $sDataField = Null])
; Parameters ....: $oImageControl       - [in/out] an object. A Image Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
;                  $sDataField          - [optional] a string value. Default is Null. The DataField value, see Remarks.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImageControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oImageControl not a Image Control.
;                  @Error 1 @Extended 3 Return 0 = $sDataField not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sDataField
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning current setting as a String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  $sDataField is not checked to make sure it exists in the referenced Database, it is the user's responsibility to do this.
;                  DataField is a String that determines the content to be shown. The entry format is either of the following:
;                  - To display the value of a column, you would call $sDataField with field:[??] where "??" represents the column's name. e.g. field:[Unique_ID].
;                  - To display the result of a function, you would call $sDataField with rpt:[??] where "??" represents the function name. e.g. rpt:[MaximumUnique_IDReport].
;                  - According to the "XReportControlModel" documentation, the following expression is also acceptable: rpt:24+24-47.
; Related .......: _LOBase_ReportConImageConGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConImageConData(ByRef $oImageControl, $sDataField = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oImageControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOBase_ReportConIdentify($oImageControl) <> $LOB_REP_CON_TYPE_IMAGE_CONTROL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sDataField) Then

		Return SetError($__LO_STATUS_SUCCESS, 1, $oImageControl.DataField())
	EndIf

	If Not IsString($sDataField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oImageControl.DataField = $sDataField
	$iError = ($oImageControl.DataField() = $sDataField) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportConImageConData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConImageConGeneral
; Description ...: Set or retrieve general Image control properties.
; Syntax ........: _LOBase_ReportConImageConGeneral(ByRef $oImageControl[, $sName = Null[, $bPreserveAsLink = Null[, $sCondPrint = Null[, $bPrintRep = Null[, $bPrintRepOnGroup = Null[, $iBackColor = Null[, $iVertAlign = Null[, $sGraphics = Null[, $iScale = Null]]]]]]]]])
; Parameters ....: $oImageControl       - [in/out] an object. A Image Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The control name.
;                  $bPreserveAsLink     - [optional] a boolean value. Default is Null. If True, the image inserted will be linked instead of embedded.
;                  $sCondPrint          - [optional] a string value. Default is Null. The conditional print expression, prefixed by "rpt:".
;                  $bPrintRep           - [optional] a boolean value. Default is Null. If True, repeated values will be printed.
;                  $bPrintRepOnGroup    - [optional] a boolean value. Default is Null. If True, repeated values will be printed on group change.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF to set Background color to default / Background Transparent = True.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOB_ALIGN_VERT_* as defined in LibreOfficeBase_Constants.au3.
;                  $sGraphics           - [optional] a string value. Default is Null. The path to an Image file.
;                  $iScale              - [optional] an integer value (0-2). Default is Null. How to scale the image to fit the button. See Constants $LOB_REP_CON_IMG_BTN_SCALE_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oImageControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oImageControl not an Image Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $bPreserveAsLink not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $sCondPrint not a String.
;                  @Error 1 @Extended 6 Return 0 = $bPrintRep not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bPrintRepOnGroup not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 9 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOB_ALIGN_VERT_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $sGraphics not a String.
;                  @Error 1 @Extended 11 Return 0 = $iScale not an Integer, less than 0 or greater than 2. See Constants $LOB_FORM_CONTROL_IMG_BTN_SCALE_* as defined in LibreOfficeBase_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $bPreserveAsLink
;                  |                               4 = Error setting $sCondPrint
;                  |                               8 = Error setting $bPrintRep
;                  |                               16 = Error setting $bPrintRepOnGroup
;                  |                               32 = Error setting $iBackColor
;                  |                               64 = Error setting $iVertAlign
;                  |                               128 = Error setting $sGraphics
;                  |                               256 = Error setting $iScale
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 9 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  I could not find a property to set the TextDirection or Visible settings.
;                  Background Transparent is set automatically based on the value set for Background color. Set Background color to $LO_COLOR_OFF to set Background Transparent to True.
; Related .......: _LOBase_ReportConImageConData, _LO_ConvertColorToLong, _LO_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConImageConGeneral(ByRef $oImageControl, $sName = Null, $bPreserveAsLink = Null, $sCondPrint = Null, $bPrintRep = Null, $bPrintRepOnGroup = Null, $iBackColor = Null, $iVertAlign = Null, $sGraphics = Null, $iScale = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[9]

	If Not IsObj($oImageControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOBase_ReportConIdentify($oImageControl) <> $LOB_REP_CON_TYPE_IMAGE_CONTROL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $bPreserveAsLink, $sCondPrint, $bPrintRep, $bPrintRepOnGroup, $iBackColor, $iVertAlign, $sGraphics, $iScale) Then
		__LO_ArrayFill($avControl, $oImageControl.Name(), $oImageControl.PreserveIRI(), $oImageControl.ConditionalPrintExpression(), $oImageControl.PrintRepeatedValues(), _
				$oImageControl.PrintWhenGroupChange(), $oImageControl.ControlBackground(), $oImageControl.VerticalAlign(), _
				_LO_PathConvert($oImageControl.ImageURL(), $LO_PATHCONV_PCPATH_RETURN), $oImageControl.ScaleMode())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oImageControl.Name = $sName
		$iError = ($oImageControl.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bPreserveAsLink <> Null) Then
		If Not IsBool($bPreserveAsLink) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oImageControl.PreserveIRI = $bPreserveAsLink
		$iError = ($oImageControl.PreserveIRI() = $bPreserveAsLink) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sCondPrint <> Null) Then
		If Not IsString($sCondPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oImageControl.ConditionalPrintExpression = $sCondPrint
		$iError = ($oImageControl.ConditionalPrintExpression() = $sCondPrint) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bPrintRep <> Null) Then
		If Not IsBool($bPrintRep) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oImageControl.PrintRepeatedValues = $bPrintRep
		$iError = ($oImageControl.PrintRepeatedValues() = $bPrintRep) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bPrintRepOnGroup <> Null) Then
		If Not IsBool($bPrintRepOnGroup) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oImageControl.PrintWhenGroupChange = $bPrintRepOnGroup
		$iError = ($oImageControl.PrintWhenGroupChange = $bPrintRepOnGroup) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oImageControl.ControlBackground = $iBackColor
		$iError = ($oImageControl.ControlBackground() = $iBackColor) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOB_ALIGN_VERT_TOP, $LOB_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oImageControl.VerticalAlign = $iVertAlign
		$iError = ($oImageControl.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($sGraphics <> Null) Then
		If Not IsString($sGraphics) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oImageControl.ImageURL = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)
		$iError = ($oImageControl.ImageURL() = _LO_PathConvert($sGraphics, $LO_PATHCONV_OFFICE_RETURN)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iScale <> Null) Then
		If Not __LO_IntIsBetween($iScale, $LOB_REP_CON_IMG_BTN_SCALE_NONE, $LOB_REP_CON_IMG_BTN_SCALE_FIT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oImageControl.ScaleMode = $iScale
		$iError = ($oImageControl.ScaleMode() = $iScale) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportConImageConGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConInsert
; Description ...: Insert a control into a report section.
; Syntax ........: _LOBase_ReportConInsert(ByRef $oSection, $iControl, $iX, $iY, $iWidth, $iHeight[, $sName = ""])
; Parameters ....: $oSection            - [in/out] an object. A section object returned by a previous _LOBase_ReportSectionGetObj, _LOBase_ReportGroupAdd, or _LOBase_ReportGroupGetByIndex function.
;                  $iControl            - an integer value (1-32). The control type to insert. See Constants $LOB_REP_CON_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  $iX                  - an integer value.The X Coordinate, in Hundredths of a Millimeter (HMM).
;                  $iY                  - an integer value. The Y Coordinate, in Hundredths of a Millimeter (HMM).
;                  $iWidth              - an integer value. The Width of the control, in Hundredths of a Millimeter (HMM).
;                  $iHeight             - an integer value. The Height of the control, in Hundredths of a Millimeter (HMM).
;                  $sName               - [optional] a string value. Default is "". The name of the control, if called with "", a name is automatically given it.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oSection not a Section Object.
;                  @Error 1 @Extended 3 Return 0 = $iControl not an Integer, less than 1 or greater than 32. See Constants $LOB_REP_CON_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $sName not a String.
;                  @Error 1 @Extended 9 Return 0 = Can't insert a chart.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create the Control.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Section parent document Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Control Service name.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Control Size Structure.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Control Position Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Control was inserted successfully, returning its object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: I presently do not know how to insert a Chart successfully.
;                  A TextBox in the L.O. Report UI is really just a formatted field, as are Page numbers, Date/ Time fields, with varying values set in DataField.
;                  A Date/Time field have either field:[TIMEVALUE(NOW())] [A Time field] or field:[TODAY()] [A Date field] set as the DataField value.
;                  A Page number field has either field:["Page " & PageNumber() & " of " & PageCount()] [A Page of pages field]; or field:["Page " & PageNumber()] [A Page field].
;                  See further note in FormattedFieldGeneral function.
;                  A Horizontal or Vertical line is a Fixed line with either Horizontal or Vertical property set using LOBase_ReportConLineGeneral function.
; Related .......: _LOBase_ReportConsGetList, _LOBase_ReportConDelete, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConInsert(ByRef $oSection, $iControl, $iX, $iY, $iWidth, $iHeight, $sName = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oControl, $oReportDoc
	Local $sControl
	Local $tPos, $tSize
	Local $iCount = 0

	If Not IsObj($oSection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSection.supportsService("com.sun.star.report.Section") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iControl, $LOB_REP_CON_TYPE_CHART, $LOB_REP_CON_TYPE_TEXT_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	If ($sName = "") Then
		Do
			$iCount += 1
			For $i = 0 To $oSection.Count() - 1
				If ($oSection.getByIndex($i).Name() = "AU3_RPT_CTRL_" & $iCount) Then ExitLoop
				Sleep((IsInt(($i / $__LOBCONST_SLEEP_DIV))) ? (10) : (0))
			Next
		Until ($i >= $oSection.Count())
	EndIf

	$oReportDoc = $oSection.ReportDefinition()
	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iControl = $LOB_REP_CON_TYPE_TEXT_BOX) Then $iControl = $LOB_REP_CON_TYPE_FORMATTED_FIELD ; Can't insert a Text Box, so exchange for Formatted Field like L.O. Base does.
	If ($iControl = $LOB_REP_CON_TYPE_CHART) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0) ; Can't insert a chart, Chart inserts as blank?? Skip charts??.

	$sControl = __LOBase_ReportConIdentify(Null, $iControl)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oControl = $oReportDoc.createInstance($sControl)
	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oControl.Name = $sName

	$tSize = $oControl.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oControl.Size = $tSize

	$oSection.Add($oControl)

	; Have to set Position after insertion.
	$tPos = $oControl.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tPos.X = $iX
	$tPos.Y = $iY

	$oControl.Position = $tPos

	Return SetError($__LO_STATUS_SUCCESS, 0, $oControl)
EndFunc   ;==>_LOBase_ReportConInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConLabelGeneral
; Description ...: Set or Retrieve general Label control settings.
; Syntax ........: _LOBase_ReportConLabelGeneral(ByRef $oLabel[, $sName = Null[, $sLabel = Null[, $sCondPrint = Null[, $bPrintRep = Null[, $bPrintRepOnGroup = Null[, $iBackColor = Null[, $mFont = Null[, $iAlign = Null[, $iVertAlign = Null]]]]]]]]])
; Parameters ....: $oLabel              - [in/out] an object. A Label Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The name of the Label control.
;                  $sLabel              - [optional] a string value. Default is Null. The Label of the control.
;                  $sCondPrint          - [optional] a string value. Default is Null. The conditional print expression, prefixed by "rpt:".
;                  $bPrintRep           - [optional] a boolean value. Default is Null. If True, repeated values will be printed.
;                  $bPrintRepOnGroup    - [optional] a boolean value. Default is Null. If True, repeated values will be printed on group change.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF to set Background color to default / Background Transparent = True.
;                  $mFont               - [optional] a map. Default is Null. A Font descriptor Map returned by a previous _LOBase_FontDescCreate or _LOBase_FontDescEdit function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The Horizontal alignment of the text. See Constants $LOB_TXT_ALIGN_HORI_* as defined in LibreOfficeBase_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOB_ALIGN_VERT_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oLabel not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oLabel not a Label Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLabel not a String.
;                  @Error 1 @Extended 5 Return 0 = $sCondPrint not a String.
;                  @Error 1 @Extended 6 Return 0 = $bPrintRep not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bPrintRepOnGroup not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 9 Return 0 = $mFont not a Map.
;                  @Error 1 @Extended 10 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants $LOB_TXT_ALIGN_HORI_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOB_ALIGN_VERT_* as defined in LibreOfficeBase_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $sLabel
;                  |                               4 = Error setting $sCondPrint
;                  |                               8 = Error setting $bPrintRep
;                  |                               16 = Error setting $bPrintRepOnGroup
;                  |                               32 = Error setting $iBackColor
;                  |                               64 = Error setting $mFont
;                  |                               128 = Error setting $iAlign
;                  |                               256 = Error setting $iVertAlign
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 9 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  I could not find a property to set the TextDirection or Visible settings.
;                  Background Transparent is set automatically based on the value set for Background color. Set Background color to $LO_COLOR_OFF to set Background Transparent to True.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConLabelGeneral(ByRef $oLabel, $sName = Null, $sLabel = Null, $sCondPrint = Null, $bPrintRep = Null, $bPrintRepOnGroup = Null, $iBackColor = Null, $mFont = Null, $iAlign = Null, $iVertAlign = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[9]

	If Not IsObj($oLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOBase_ReportConIdentify($oLabel) <> $LOB_REP_CON_TYPE_LABEL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $sLabel, $sCondPrint, $bPrintRep, $bPrintRepOnGroup, $iBackColor, $mFont, $iAlign, $iVertAlign) Then
		__LO_ArrayFill($avControl, $oLabel.Name(), $oLabel.Label(), $oLabel.ConditionalPrintExpression(), $oLabel.PrintRepeatedValues(), _
				$oLabel.PrintWhenGroupChange(), $oLabel.ControlBackground(), __LOBase_ReportConSetGetFontDesc($oLabel), $oLabel.ParaAdjust(), _
				$oLabel.VerticalAlign())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oLabel.Name = $sName
		$iError = ($oLabel.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oLabel.Label = $sLabel
		$iError = ($oLabel.Label() = $sLabel) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sCondPrint <> Null) Then
		If Not IsString($sCondPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oLabel.ConditionalPrintExpression = $sCondPrint
		$iError = ($oLabel.ConditionalPrintExpression() = $sCondPrint) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bPrintRep <> Null) Then
		If Not IsBool($bPrintRep) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oLabel.PrintRepeatedValues = $bPrintRep
		$iError = ($oLabel.PrintRepeatedValues() = $bPrintRep) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bPrintRepOnGroup <> Null) Then
		If Not IsBool($bPrintRepOnGroup) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oLabel.PrintWhenGroupChange = $bPrintRepOnGroup
		$iError = ($oLabel.PrintWhenGroupChange = $bPrintRepOnGroup) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oLabel.ControlBackground = $iBackColor
		$iError = ($oLabel.ControlBackground() = $iBackColor) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($mFont <> Null) Then
		If Not IsMap($mFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		__LOBase_ReportConSetGetFontDesc($oLabel, $mFont)
		$iError = (@error = 0) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOB_TXT_ALIGN_HORI_LEFT, $LOB_TXT_ALIGN_HORI_CENTER) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oLabel.ParaAdjust = $iAlign
		$iError = ($oLabel.ParaAdjust() = $iAlign) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOB_ALIGN_VERT_TOP, $LOB_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oLabel.VerticalAlign = $iVertAlign
		$iError = ($oLabel.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportConLabelGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConLineGeneral
; Description ...: Set or Retrieve general Line control settings.
; Syntax ........: _LOBase_ReportConLineGeneral(ByRef $oLine[, $sName = Null[, $iVertAlign = Null[, $iOrient = Null]]])
; Parameters ....: $oLine               - [in/out] an object. A Line Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
;                  $sName               - [optional] a string value. Default is Null. The name of the control.
;                  $iVertAlign          - [optional] an integer value (0-2). Default is Null. The Vertical alignment of the text. See Constants $LOB_ALIGN_VERT_* as defined in LibreOfficeBase_Constants.au3.
;                  $iOrient             - [optional] an integer value (0-1). Default is Null. The orientation of the line. See Constants $LOB_REP_CON_LINE_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oLabel not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oLabel not a Label Control.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iVertAlign not an Integer, less than 0 or greater than 2. See Constants $LOB_ALIGN_VERT_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iOrient not an Integer, less than 0 or greater than 1. See Constants $LOB_REP_CON_LINE_* as defined in LibreOfficeBase_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control type.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $iVertAlign
;                  |                               4 = Error setting $iOrient
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  I could not find a property to set "Visible" setting.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConLineGeneral(ByRef $oLine, $sName = Null, $iVertAlign = Null, $iOrient = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avControl[3]

	If Not IsObj($oLine) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (__LOBase_ReportConIdentify($oLine) <> $LOB_REP_CON_TYPE_LINE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sName, $iVertAlign, $iOrient) Then
		__LO_ArrayFill($avControl, $oLine.Name(), $oLine.VerticalAlign(), $oLine.Orientation())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avControl)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oLine.Name = $sName
		$iError = ($oLine.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOB_ALIGN_VERT_TOP, $LOB_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oLine.VerticalAlign = $iVertAlign
		$iError = ($oLine.VerticalAlign() = $iVertAlign) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iOrient <> Null) Then
		If Not __LO_IntIsBetween($iOrient, $LOB_REP_CON_LINE_HORI, $LOB_REP_CON_LINE_VERT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oLine.Orientation = $iOrient
		$iError = ($oLine.Orientation() = $iOrient) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportConLineGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConnect
; Description ...: Retrieve an Object for the currently open Report or Reports.
; Syntax ........: _LOBase_ReportConnect([$bConnectCurrent = True])
; Parameters ....: $bConnectCurrent     - [optional] a boolean value. Default is True. If True, Returns an Object for the last active Report. Else an array of all Open Reports. See Remarks.
; Return values .: Success: Object or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bConnectCurrent not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.ServiceManager Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create com.sun.star.frame.Desktop Object.
;                  @Error 2 @Extended 3 Return 0 = Failed to create enumeration of open Documents.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = No LibreOffice windows are open.
;                  @Error 3 @Extended 2 Return 0 = Current LibreOffice window is not a Report Document.
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Success. Connected to the currently active window, returning the Report Document Object. Report is in Read-Only viewing mode.
;                  @Error 0 @Extended 2 Return Object = Success. Connected to the currently active window, returning the Report Document Object. Report is in Design mode.
;                  @Error 0 @Extended ? Return Array = Success. Returning a Three columned Array with all open Report Documents. @Extended is set to the number of results. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Returned array when connecting to all open Report Documents returns an array with Three columns per result. ($aArray[0][3]). Each result is stored in a separate row;
;                  Row 1, Column 0 contain the Object for that document. e.g. $aArray[0][0] = $oDoc
;                  Row 1, Column 1 contains the Document's full title with extension and the Report Name, separated by a colon. e.g. $aArray[0][1] = "Testing.odb : Report1"
;                  Row 1, Column 2 contains a Boolean value whether the Report is in Design mode (True) or not.
;                  Row 2, Column 0 contain the Object for the next document. And so on. e.g. $aArray[1][0] = $oDoc2
; Related .......: _LOBase_ReportOpen, _LOBase_ReportClose
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConnect($bConnectCurrent = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local $aoConnectAll[0][3]
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop
	Local $sReportViewServiceName = "com.sun.star.text.TextDocument", $sReportDesignServiceName = "com.sun.star.report.ReportDefinition"

	If Not IsBool($bConnectCurrent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	If Not $oDesktop.getComponents.hasElements() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; no L.O open

	$oEnumDoc = $oDesktop.getComponents.createEnumeration()
	If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If $bConnectCurrent Then
		$oDoc = $oDesktop.currentComponent()
		If ($oDoc.supportsService($sReportViewServiceName) And $oDoc.isReadOnly() And Not (IsObj($oDoc.Parent()))) Then ; View only Report Doc.

			Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc)

		ElseIf $oDoc.supportsService($sReportDesignServiceName) Then ; Report Doc in Design mode.

			Return SetError($__LO_STATUS_SUCCESS, 2, $oDoc)

		Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndIf

	Else
		ReDim $aoConnectAll[1][3]
		$iCount = 0
		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService($sReportDesignServiceName) _ ; Report Doc in Design mode.
					Or ($oDoc.supportsService($sReportViewServiceName) And $oDoc.isReadOnly() And Not (IsObj($oDoc.Parent()))) Then ; If Parent is not present and document is Read-Only, it should be a Database Report.

				ReDim $aoConnectAll[$iCount + 1][3]
				$aoConnectAll[$iCount][0] = $oDoc
				$aoConnectAll[$iCount][1] = $oDoc.Title()
				$aoConnectAll[$iCount][2] = ($oDoc.supportsService($sReportDesignServiceName)) ? (True) : (False)
				$iCount += 1
			EndIf
			Sleep(10)
		WEnd

		Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoConnectAll)
	EndIf
EndFunc   ;==>_LOBase_ReportConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConPosition
; Description ...: Set or Retrieve the Control's position settings.
; Syntax ........: _LOBase_ReportConPosition(ByRef $oControl[, $iX = Null[, $iY = Null]])
; Parameters ....: $oControl            - [in/out] an object. A Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
;                  $iX                  - [optional] an integer value. Default is Null. The X position from the insertion point, in Hundredths of a Millimeter (HMM).
;                  $iY                  - [optional] an integer value. Default is Null. The Y position from the insertion point, in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iY not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Control's Position Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iX
;                  |                               2 = Error setting $iY
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LO_UnitConvert, _LOBase_ReportConSize
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConPosition(ByRef $oControl, $iX = Null, $iY = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPosition[2]
	Local $tPos

	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tPos = $oControl.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iX, $iY) Then
		__LO_ArrayFill($avPosition, $tPos.X(), $tPos.Y())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPosition)
	EndIf

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

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportConPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConsGetList
; Description ...: Retrieve an array of Control Objects contained in a Report's Section.
; Syntax ........: _LOBase_ReportConsGetList(ByRef $oSection[, $iType = $LOB_REP_CON_TYPE_ALL])
; Parameters ....: $oSection            - [in/out] an object. A section object returned by a previous _LOBase_ReportSectionGetObj, _LOBase_ReportGroupAdd, or _LOBase_ReportGroupGetByIndex function.
;                  $iType               - [optional] an integer value (1-63). Default is $LOB_REP_CON_TYPE_ALL. The type of control(s) to return in the array. Can be BitOr'd together. See Constants $LOB_REP_CON_TYPE_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oSection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Called Object not a Section Object.
;                  @Error 1 @Extended 3 Return 0 = $iType not an Integer, less than 1 or greater than 63. See Constants $LOB_REP_CON_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Control Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify Control type.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning a 2D array of Control Objects in the first column, and the type of Control in the second column, corresponding to the Constants $LOB_REP_CON_TYPE_* as defined in LibreOfficeBase_Constants.au3
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_ReportConDelete, _LOBase_ReportConInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConsGetList(ByRef $oSection, $iType = $LOB_REP_CON_TYPE_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aoControls[0][2]
	Local $oControl
	Local $iCount = 0, $iControlType

	If Not IsObj($oSection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSection.supportsService("com.sun.star.report.Section") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iType, $LOB_REP_CON_TYPE_CHART, $LOB_REP_CON_TYPE_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	ReDim $aoControls[$oSection.Count()][2]

	For $i = 0 To $oSection.Count() - 1
		$oControl = $oSection.getByIndex($i)
		If Not IsObj($oControl) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$iControlType = __LOBase_ReportConIdentify($oControl)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		If BitAND($iType, $iControlType) Then
			$aoControls[$iCount][0] = $oControl
			$aoControls[$iCount][1] = $iControlType
			$iCount += 1
		EndIf

		Sleep((IsInt(($i / $__LOBCONST_SLEEP_DIV))) ? (10) : (0))
	Next

	ReDim $aoControls[$iCount][2]

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoControls)
EndFunc   ;==>_LOBase_ReportConsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportConSize
; Description ...: Set or Retrieve Control Size or related settings.
; Syntax ........: _LOBase_ReportConSize(ByRef $oControl[, $iWidth = Null[, $iHeight = Null[, $bAutoGrow = Null]]])
; Parameters ....: $oControl            - [in/out] an object. A Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the Shape, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the Shape, in Hundredths of a Millimeter (HMM). Min. 51.
;                  $bAutoGrow           - [optional] a boolean value. Default is Null. If True, the control's size will automatically adjust to fit content.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer, or less than 51.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer, or less than 51.
;                  @Error 1 @Extended 4 Return 0 = $bAutoGrow not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Control Size Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iHeight
;                  |                               4 = Error setting $bAutoGrow
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_UnitConvert, _LOBase_ReportConPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportConSize(ByRef $oControl, $iWidth = Null, $iHeight = Null, $bAutoGrow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSize[3]
	Local $tSize

	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tSize = $oControl.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iHeight, $bAutoGrow) Then
		__LO_ArrayFill($avSize, $tSize.Width(), $tSize.Height(), $oControl.AutoGrow())

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

	If ($bAutoGrow <> Null) Then
		If Not IsBool($bAutoGrow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oControl.AutoGrow = $bAutoGrow
		$iError = ($oControl.AutoGrow() = $bAutoGrow) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportConSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportCopy
; Description ...: Create a copy of an existing Report.
; Syntax ........: _LOBase_ReportCopy(ByRef $oConnection, $sInputReport, $sOutputReport)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sInputReport        - a string value. The Name of the Report to Copy. Also the Sub-directory the Report is in. See Remarks.
;                  $sOutputReport       - a string value. The Name of the Report to Create. Also the Sub-directory to place the Report in. See Remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sInputReport not a String.
;                  @Error 1 @Extended 3 Return 0 = $sOutputReport not a String.
;                  @Error 1 @Extended 4 Return 0 = Requested report not found.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sInputReport not a Report.
;                  @Error 1 @Extended 6 Return 0 = Folder name called in $sOutputReport not found.
;                  @Error 1 @Extended 7 Return 0 = Report already exists with called name in Destination.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.sdb.DocumentDefinition" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Report Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Destination Report name.
;                  @Error 3 @Extended 5 Return 0 = Failed to insert copied Report.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Copied report successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To copy a Report located inside a folder, the Report name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to copy ReportXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sInputReport with the following path: Folder1/Folder2/Folder3/ReportXYZ.
;                  To create a Report inside a folder, the Report name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to create ReportXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sOutputReport with the following path: Folder1/Folder2/Folder3/ReportXYZ.
;                  If only a name is called in $sOutputReport, the Report will be created in the main directory, i.e. not inside of any folders.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportCopy(ByRef $oConnection, $sInputReport, $sOutputReport)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oRepDef, $oDocDef
	Local $aArgs[3]
	Local $sDestReport

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sInputReport) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sOutputReport) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSource = $oConnection.Parent.DatabaseDocument.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oSource.hasByHierarchicalName($sInputReport) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oRepDef = $oSource.getByHierarchicalName($sInputReport)
	If Not IsObj($oRepDef) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If Not $oRepDef.supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If StringInStr($sOutputReport, "/") And Not $oSource.hasByHierarchicalName(StringLeft($sOutputReport, StringInStr($sOutputReport, "/", 0, -1) - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If $oSource.hasByHierarchicalName($sOutputReport) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$sDestReport = StringTrimLeft($sOutputReport, StringInStr($sOutputReport, "/", 0, -1))
	If Not IsString($sDestReport) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$aArgs[0] = __LO_SetPropertyValue("Name", $sDestReport)
	$aArgs[1] = __LO_SetPropertyValue("ActiveConnection", $oConnection)
	$aArgs[2] = __LO_SetPropertyValue("EmbeddedObject", $oRepDef)

	$oDocDef = $oSource.createInstanceWithArguments("com.sun.star.sdb.DocumentDefinition", $aArgs)
	If Not IsObj($oDocDef) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oSource.insertByHierarchicalName($sOutputReport, $oDocDef)
	If Not $oSource.hasByHierarchicalName($sOutputReport) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportCreate
; Description ...: Create and Insert a new Report Document into a Base Document.
; Syntax ........: _LOBase_ReportCreate(ByRef $oConnection, $sReport[, $bOpen = False[, $bHidden = False]])
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sReport             - a string value. The Name of the Report to Create. Also the Sub-directory to place the Report in. See Remarks.
;                  $bOpen               - [optional] a boolean value. Default is False. If True, the new Report will be opened in Design mode.
;                  $bHidden             - [optional] a boolean value. Default is False. If True, the Report will be invisible when opened.
; Return values .: Success: 1 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sReport not a String.
;                  @Error 1 @Extended 3 Return 0 = $bOpen not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = Folder or Sub-Folder not found.
;                  @Error 1 @Extended 6 Return 0 = Name called in $sReport already exists in Folder.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create com.sun.star.sdb.DocumentDefinition Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Document URL.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report name.
;                  @Error 3 @Extended 5 Return 0 = Failed to insert new Report into Base Document.
;                  @Error 3 @Extended 6 Return 0 = Failed to open new Report Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. New Report was successfully inserted.
;                  @Error 0 @Extended 1 Return Object = Success. Returning opened Report Document's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To create a Report inside a folder, the Report name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to create ReportXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sReport with the following path: Folder1/Folder2/Folder3/ReportXYZ.
;                  When created, the report will not have a Data source set, so it will not be able to be opened in viewing mode, only in Design mode.
;                  When created, the report will have neither Page Header nor Page Footer enabled.
;                  Thanks to sokol92 on the LibreOffice forum for this method. https://ask.libreoffice.org/t/create-a-new-report-document-using-a-macro/123584/16?u=donh1
; Related .......: _LOBase_ReportDelete, _LOBase_ReportCopy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportCreate(ByRef $oConnection, $sReport, $bOpen = False, $bHidden = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oReportDoc, $oDocDef
	Local $aArgs[3]
	Local $sDocURL, $sReportName

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sReport) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOpen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSource = $oConnection.Parent.DatabaseDocument.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$sDocURL = $oSource.Parent.URL()
	If Not IsString($sDocURL) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If StringInStr($sReport, "/") And Not $oSource.hasByHierarchicalName(StringLeft($sReport, StringInStr($sReport, "/", 0, -1) - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If $oSource.hasByHierarchicalName($sReport) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$sReportName = StringTrimLeft($sReport, StringInStr($sReport, "/", 0, -1))
	If Not IsString($sReportName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$aArgs[0] = __LO_SetPropertyValue("Name", $sReportName)
	$aArgs[1] = __LO_SetPropertyValue("Parent", $oSource)
	$aArgs[2] = __LO_SetPropertyValue("URL", _LO_PathConvert($sDocURL, $LO_PATHCONV_OFFICE_RETURN))

	$oDocDef = $oSource.createInstanceWithArguments("com.sun.star.sdb.DocumentDefinition", $aArgs)
	If Not IsObj($oDocDef) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oSource.insertByHierarchicalName($sReport, $oDocDef)

	If Not $oSource.hasByHierarchicalName($sReport) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	If $bOpen Then
		If Not $oSource.Parent.CurrentController.isConnected() Then $oSource.Parent.CurrentController.connect()

		ReDim $aArgs[1]
		$aArgs[0] = __LO_SetPropertyValue("Hidden", $bHidden)

		$oReportDoc = $oSource.Parent.CurrentController.loadComponentWithArguments($LOB_SUB_COMP_TYPE_REPORT, $sReport, True, $aArgs)
		If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $oReportDoc)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportData
; Description ...: Set or Retrieve Data related properties for a Report Document.
; Syntax ........: _LOBase_ReportData(ByRef $oReportDoc[, $iContentType = Null[, $sContent = Null[, $bAnalyzeSQL = Null[, $sFilter = Null[, $iReportOutput = Null[, $bSuppress = Null]]]]]])
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $iContentType        - [optional] an integer value (0-2). Default is Null. The Type of data source for the Report. See Constants, $LOB_REP_CONTENT_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  $sContent            - [optional] a string value. Default is Null. The Content to be used for the Report, either a Table or Query name, or an SQL statement.
;                  $bAnalyzeSQL         - [optional] a boolean value. Default is Null. If True, SQL commands will be analyzed by LibreOffice.
;                  $sFilter             - [optional] a string value. Default is Null. The SQL filter command.
;                  $iReportOutput       - [optional] an integer value (1-2). Default is Null. The type of output document when the Report is executed. See Constants, $LOB_REP_OUTPUT_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  $bSuppress           - [optional] a boolean value. Default is Null. If True, the "Add a Field" dialog will be suppressed from coming up. See remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $iContentType not an Integer, less than 0 or greater than 2. See Constants, $LOB_REP_CONTENT_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $sContent not a String.
;                  @Error 1 @Extended 5 Return 0 = $bAnalyzeSQL not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $sFilter not a String.
;                  @Error 1 @Extended 7 Return 0 = $iReportOutput not an Integer, less than 1 or greater than 2. See Constants, $LOB_REP_OUTPUT_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $bSuppress not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iContentType
;                  |                               2 = Error setting $sContent
;                  |                               4 = Error setting $bAnalyzeSQL
;                  |                               8 = Error setting $sFilter
;                  |                               16 = Error setting $iReportOutput
;                  |                               32 = Error setting $bSuppress
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Modifying $iContentType and $sContent  will open the "Add a Field" dialog unless it is suppressed.
;                  When $bSuppress is True, changing $iContentType and $sContent, either in the UI or via AutoIt, will not activate the "Add a Field" dialog, until the report is re-opened again, or $bSuppress is called with False again.
;                  Setting $bSuppress to False will activate the "Add a Field" dialog, regardless if any settings are changed.
; Related .......: _LOBase_ReportGeneral
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportData(ByRef $oReportDoc, $iContentType = Null, $sContent = Null, $bAnalyzeSQL = Null, $sFilter = Null, $iReportOutput = Null, $bSuppress = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avReport[6]
	Local Const $__LOB_REP_OUTPUT_TEXT_DOC = "application/vnd.oasis.opendocument.text", $__LOB_REP_OUTPUT_SPREADSHEET_DOC = "application/vnd.oasis.opendocument.spreadsheet"

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iContentType, $sContent, $bAnalyzeSQL, $sFilter, $iReportOutput, $bSuppress) Then
		__LO_ArrayFill($avReport, $oReportDoc.CommandType(), $oReportDoc.Command(), $oReportDoc.EscapeProcessing(), $oReportDoc.Filter(), _
				($oReportDoc.MimeType() = $__LOB_REP_OUTPUT_TEXT_DOC) ? ($LOB_REP_OUTPUT_TYPE_TEXT) : (($oReportDoc.MimeType() = $__LOB_REP_OUTPUT_SPREADSHEET_DOC) ? ($LOB_REP_OUTPUT_TYPE_SPREADSHEET) : ($LOB_REP_OUTPUT_TYPE_UNKNOWN)), _
				($oReportDoc.CurrentController.Mode() = "normal") ? (False) : (True))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avReport)
	EndIf

	If ($iContentType <> Null) Then
		If Not __LO_IntIsBetween($iContentType, $LOB_REP_CONTENT_TYPE_TABLE, $LOB_REP_CONTENT_TYPE_SQL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If ($bSuppress = True) Then $oReportDoc.CurrentController.Mode = "remote"
		$oReportDoc.CommandType = $iContentType
		$iError = ($oReportDoc.CommandType() = $iContentType) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If ($bSuppress = True) Then $oReportDoc.CurrentController.Mode = "remote"
		$oReportDoc.Command = $sContent
		$iError = ($oReportDoc.Command() = $sContent) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bAnalyzeSQL <> Null) Then
		If Not IsBool($bAnalyzeSQL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oReportDoc.EscapeProcessing = $bAnalyzeSQL
		$iError = ($oReportDoc.EscapeProcessing() = $bAnalyzeSQL) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($sFilter <> Null) Then
		If Not IsString($sFilter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oReportDoc.Filter = $sFilter
		$iError = ($oReportDoc.Filter() = $sFilter) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iReportOutput <> Null) Then
		If Not __LO_IntIsBetween($iReportOutput, $LOB_REP_OUTPUT_TYPE_TEXT, $LOB_REP_OUTPUT_TYPE_SPREADSHEET) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oReportDoc.MimeType = ($iReportOutput = $LOB_REP_OUTPUT_TYPE_TEXT) ? ($__LOB_REP_OUTPUT_TEXT_DOC) : ($__LOB_REP_OUTPUT_SPREADSHEET_DOC)
		$iError = ($oReportDoc.MimeType() = ($iReportOutput = $LOB_REP_OUTPUT_TYPE_TEXT) ? ($__LOB_REP_OUTPUT_TEXT_DOC) : ($__LOB_REP_OUTPUT_SPREADSHEET_DOC)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bSuppress <> Null) Then
		If Not IsBool($bSuppress) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oReportDoc.CurrentController.Mode = ($bSuppress) ? ("remote") : ("normal") ; Remote prevents the Add Field dialog from coming up when changing "Command" and "CommandType". Normal behaves as normal.
		$iError = ($oReportDoc.CurrentController.Mode = ($bSuppress) ? ("remote") : ("normal")) ? ($iError) : (BitOR($iError, 32)) ; Method found in "ReportBuilderImplementation.java" file, line 160, function "switchOffAddFieldWindow"
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportData

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportDelete
; Description ...: Delete a Report from a Document.
; Syntax ........: _LOBase_ReportDelete(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sName               - a string value. The Report name to Delete. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Name called in $sName not found in Folder.
;                  @Error 1 @Extended 4 Return 0 = Name called in $sName not a Report.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to delete Report.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Report was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To delete a report contained in a folder, you MUST prefix the Report name called in $sName by the folder path it is located in, separated by forward slashes (/). e.g. to delete ReportXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/ReportXYZ
; Related .......: _LOBase_ReportCopy, _LOBase_ReportsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportDelete(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oSource.hasByHierarchicalName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not $oSource.getByHierarchicalName($sName).supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oSource.removeByHierarchicalName($sName)

	If $oSource.hasByHierarchicalName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportDetail
; Description ...: Set or Retrieve Report Detail section properties.
; Syntax ........: _LOBase_ReportDetail(ByRef $oReportDoc[, $sName = Null[, $iForceNewPage = Null[, $bKeepTogether = Null[, $bVisible = Null[, $iHeight = Null[, $sCondPrint = Null[, $iBackColor = Null]]]]]]])
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $sName               - [optional] a string value. Default is Null. The name of the Section.
;                  $iForceNewPage       - [optional] an integer value (0-3). Default is Null. If and when to force a new page. See Constants, $LOB_REP_FORCE_PAGE_* as defined in LibreOfficeBase_Constants.au3.
;                  $bKeepTogether       - [optional] a boolean value. Default is Null. If True, the section should be printed on one page.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the section is visible in the Report.
;                  $iHeight             - [optional] an integer value (1753-??). Default is Null. The height of the Section, in Hundredths of a Millimeter (HMM). See remarks.
;                  $sCondPrint          - [optional] a string value. Default is Null. The Conditional Print Statement.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF to set Background color to default / Background Transparent = True.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iForceNewPage not an Integer, less than 0 or greater than 3. See Constants, $LOB_REP_FORCE_PAGE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bKeepTogether not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iHeight not an Integer, or less than 1753.
;                  @Error 1 @Extended 8 Return 0 = $sCondPrint not a String.
;                  @Error 1 @Extended 9 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $iForceNewPage
;                  |                               4 = Error setting $bKeepTogether
;                  |                               8 = Error setting $bVisible
;                  |                               16 = Error setting $iHeight
;                  |                               32 = Error setting $sCondPrint
;                  |                               64 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The minimum height of a Section is 1753 Hundredths of a Millimeter (HMM), the maximum is unknown, but I found that setting a large value tends to cause a freeze up/crash of the Report.
;                  Background Transparent is set automatically based on the value set for Background color. Set Background color to $LO_COLOR_OFF to set Background Transparent to True.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_ReportPageFooter, _LOBase_ReportPageHeader, _LOBase_ReportFooter, _LOBase_ReportHeader, _LOBase_ReportGroupFooter, _LOBase_ReportGroupHeader, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportDetail(ByRef $oReportDoc, $sName = Null, $iForceNewPage = Null, $bKeepTogether = Null, $bVisible = Null, $iHeight = Null, $sCondPrint = Null, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avProps[7]
	Local $iError = 0

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sName, $iForceNewPage, $bKeepTogether, $bVisible, $iHeight, $sCondPrint, $iBackColor) Then
		__LO_ArrayFill($avProps, $oReportDoc.Detail.Name(), $oReportDoc.Detail.ForceNewPage(), $oReportDoc.Detail.KeepTogether(), $oReportDoc.Detail.Visible(), _
				$oReportDoc.Detail.Height(), $oReportDoc.Detail.ConditionalPrintExpression(), $oReportDoc.Detail.BackColor())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avProps)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oReportDoc.Detail.Name = $sName
		$iError = ($oReportDoc.Detail.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iForceNewPage <> Null) Then
		If Not __LO_IntIsBetween($iForceNewPage, $LOB_REP_FORCE_PAGE_NONE, $LOB_REP_FORCE_PAGE_BEFORE_AFTER_SECTION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oReportDoc.Detail.ForceNewPage = $iForceNewPage
		$iError = ($oReportDoc.Detail.ForceNewPage() = $iForceNewPage) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bKeepTogether <> Null) Then
		If Not IsBool($bKeepTogether) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oReportDoc.Detail.KeepTogether = $bKeepTogether
		$iError = ($oReportDoc.Detail.KeepTogether() = $bKeepTogether) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oReportDoc.Detail.Visible = $bVisible
		$iError = ($oReportDoc.Detail.Visible() = $bVisible) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iHeight <> Null) Then
		If Not __LO_IntIsBetween($iHeight, 1753) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oReportDoc.Detail.Height = $iHeight
		$iError = (__LO_IntIsBetween($oReportDoc.Detail.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($sCondPrint <> Null) Then
		If Not IsString($sCondPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oReportDoc.Detail.ConditionalPrintExpression = $sCondPrint
		$iError = ($oReportDoc.Detail.ConditionalPrintExpression() = $sCondPrint) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oReportDoc.Detail.BackColor = $iBackColor
		$iError = ($oReportDoc.Detail.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 64))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportDetail

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormDocVisible
; Description ...: Set or retrieve the current visibility of a document.
; Syntax ........: _LOBase_FormDocVisible(ByRef $oReportDoc[, $bVisible = Null])
; Parameters ....: $oReportDoc          - [in/out] an object.  A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the document is visible.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. $bVisible successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. Returning current visibility state of the Document, True if visible, False if invisible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call $bVisible with Null to return the current visibility setting.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportDocVisible(ByRef $oReportDoc, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bVisible) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oReportDoc.CurrentController.Frame.ContainerWindow.isVisible())

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oReportDoc.CurrentController.Frame.ContainerWindow.Visible = $bVisible
	$iError = ($oReportDoc.CurrentController.Frame.ContainerWindow.isVisible() = $bVisible) ? (0) : (1)

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOBase_ReportDocVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportExists
; Description ...: Check whether a Document contains a Report by name.
; Syntax ........: _LOBase_ReportExists(ByRef $oDoc, $sName[, $bExhaustive = True])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sName               - a string value. The name of the Report to look for. See remarks.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, the search looks inside sub-folders.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended ? Return Boolean = Success. Returning a Boolean value indicating if the Document contains a Report by the called name (True) or not. If True, and $bExhaustive is True, @Extended is set to the number of times a Report with the same name is found in the Document (In sub-folders).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To narrow the search for a Report down to a specific folder, you MUST prefix the Report name called in $sName by the folder path to look in, separated by forward slashes (/). e.g. to search for ReportXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/ReportXYZ
; Related .......: _LOBase_ReportsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportExists(ByRef $oDoc, $sName, $bExhaustive = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local $bReturn = False
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iReports = 0
	Local $asNames[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Report name to search

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If $oSource.hasByName($sName) And $oSource.getByName($sName).supportsService("com.sun.star.ucb.Content") Then
		$iReports += 1
		$bReturn = True
	EndIf

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			ReDim $avFolders[1][2]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				If $avFolders[$iCount][$iObj].hasByName($sName) And $avFolders[$iCount][$iObj].getByName($sName).supportsService("com.sun.star.ucb.Content") Then
					$iReports += 1
					$bReturn = True
				EndIf

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						ReDim $avFolders[$iFolders + 1][2]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][2]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iReports, $bReturn)
EndFunc   ;==>_LOBase_ReportExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFolderCopy
; Description ...: Create a copy of an existing Folder.
; Syntax ........: _LOBase_ReportFolderCopy(ByRef $oDoc, $sInputFolder, $sOutputFolder)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sInputFolder        - a string value. The Name of the Folder to Copy. Also the Sub-directory the Folder is in. See Remarks.
;                  $sOutputFolder       - a string value. The Name of the Folder to Create. Also the Sub-directory to place the Folder in. See Remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sInputFolder not a String.
;                  @Error 1 @Extended 3 Return 0 = $sOutputFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Requested Folder not found.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sInputFolder not a Folder.
;                  @Error 1 @Extended 6 Return 0 = Folder already exists with called name in Destination.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.sdb.Reports" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Source Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Destination Folder name.
;                  @Error 3 @Extended 4 Return 0 = Failed to insert copied Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Copied Folder successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To create a Folder contained in a folder, you MUST prefix the Folder name called in $sInputFolder by the folder path it is located in, separated by forward slashes (/). e.g. to copy FolderXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sInputFolder with the following path: Folder1/Folder2/Folder3/FolderXYZ
;                  To copy a Folder contained in a folder, you MUST prefix the Folder name called in $sOutputFolder by the folder path you want it to be located in, separated by forward slashes (/). e.g. to create FolderXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sOutputFolder with the following path: Folder1/Folder2/Folder3/FolderXYZ
;                  Copying a Folder will copy all contents also.
;                  If only a name is called in $sOutputFolder, the Folder will be created in the main directory, i.e. not inside of any folders.
; Related .......: _LOBase_ReportFolderCreate, _LOBase_ReportFolderDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFolderCopy(ByRef $oDoc, $sInputFolder, $sOutputFolder)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oSourceReportFolder, $oFolder
	Local $aArgs[2]
	Local $sDestFolder

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sInputFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sOutputFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oSource.hasByHierarchicalName($sInputFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oSourceReportFolder = $oSource.getByHierarchicalName($sInputFolder)
	If Not IsObj($oSourceReportFolder) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oSourceReportFolder.supportsService("com.sun.star.sdb.Reports") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$sDestFolder = StringTrimLeft($sOutputFolder, StringInStr($sOutputFolder, "/", 0, -1))
	If Not IsString($sDestFolder) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If $oSource.hasByHierarchicalName($sOutputFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$aArgs[0] = __LO_SetPropertyValue("Name", $sDestFolder)
	$aArgs[1] = __LO_SetPropertyValue("EmbeddedObject", $oSourceReportFolder)

	$oFolder = $oSource.createInstanceWithArguments("com.sun.star.sdb.Reports", $aArgs)
	If Not IsObj($oFolder) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oSource.insertByHierarchicalName($sOutputFolder, $oFolder)
	If Not $oSource.hasByHierarchicalName($sOutputFolder) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportFolderCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFolderCreate
; Description ...: Create a Report Folder.
; Syntax ........: _LOBase_ReportFolderCreate(ByRef $oDoc, $sFolder[, $bCreateMulti = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sFolder             - a string value. The Folder name to create. Can also include the sub-folder path. See Remarks.
;                  $bCreateMulti        - [optional] a boolean value. Default is False. If True, multiple folders in a path will be created if they do not exist.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 3 Return 0 = $bCreateMulti not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = Name called in $sFolder already exists in Folder.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Folder Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to insert new Folder into Base Document.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Destination Folder Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully created the Folder(s).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To create a Folder inside a folder, the Folder name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to create FolderXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3/FolderXYZ.
; Related .......: _LOBase_ReportFolderCopy, _LOBase_ReportFolderDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFolderCreate(ByRef $oDoc, $sFolder, $bCreateMulti = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oObj
	Local $asSplit[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bCreateMulti) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oDoc.ReportDocuments.hasByHierarchicalName($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If $bCreateMulti Then
		$asSplit = StringSplit($sFolder, "/")

		For $i = 1 To $asSplit[0]
			If Not $oSource.hasByName($asSplit[$i]) Then
				$oObj = $oSource.createInstance("com.sun.star.sdb.Reports")
				If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

				$oSource.insertbyName($asSplit[$i], $oObj)

				If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				$oSource = $oSource.getByName($asSplit[$i])
				If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			Else
				$oSource = $oSource.getByName($asSplit[$i])
				If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
			EndIf
		Next

	Else
		$oObj = $oSource.createInstance("com.sun.star.sdb.Reports")
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oSource.insertByHierarchicalName($sFolder, $oObj)
	EndIf

	If Not $oDoc.ReportDocuments.hasByHierarchicalName($sFolder) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportFolderCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFolderDelete
; Description ...: Delete a Report Folder from a Document.
; Syntax ........: _LOBase_ReportFolderDelete(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sName               - a string value. The Folder name to Delete. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Name called in $sName not found in Folder.
;                  @Error 1 @Extended 4 Return 0 = Name called in $sName not a Folder.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to delete Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Folder was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To delete a Folder contained in a folder, you MUST prefix the Folder name called in $sName by the folder path it is located in, separated by forward slashes (/). e.g. to delete FolderXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/FolderXYZ
;                  Deleting a Folder will delete all contents also.
; Related .......: _LOBase_ReportFolderCreate _LOBase_ReportFolderCopy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFolderDelete(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oSource.hasByHierarchicalName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not $oSource.getByHierarchicalName($sName).supportsService("com.sun.star.sdb.Reports") Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oSource.removeByHierarchicalName($sName)

	If $oSource.hasByHierarchicalName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportFolderDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFolderExists
; Description ...: Check whether a Document contains a Report Folder by name.
; Syntax ........: _LOBase_ReportFolderExists(ByRef $oDoc, $sName[, $bExhaustive = True])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sName               - a string value. The name of the Folder to look for.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, the search looks inside sub-folders.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended ? Return Boolean = Success. Returning a Boolean value indicating if the Document contains a Folder by the called name (True) or not. If True, and $bExhaustive is True, @Extended is set to the number of times a Folder with the same name is found in the Document (In sub-folders).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To narrow the search for a Folder down to a specific folder, you MUST prefix the Folder name called in $sName by the folder path to look in, separated by forward slashes (/). e.g. to search for FolderXYZ located in folder3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/FolderXYZ
; Related .......: _LOBase_ReportFoldersGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFolderExists(ByRef $oDoc, $sName, $bExhaustive = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local $bReturn = False
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iResults = 0
	Local $asNames[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asSplit = StringSplit($sName, "/")

	For $i = 1 To $asSplit[0] - 1
		If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oSource = $oSource.getByName($asSplit[$i])
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	Next

	$sName = $asSplit[$asSplit[0]] ; Last element of Array will be the Folder name to search

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If $oSource.hasByName($sName) And $oSource.getByName($sName).supportsService("com.sun.star.sdb.Reports") Then
		$iResults += 1
		$bReturn = True
	EndIf

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			ReDim $avFolders[1][2]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				If $avFolders[$iCount][$iObj].hasByName($sName) And $avFolders[$iCount][$iObj].getByName($sName).supportsService("com.sun.star.sdb.Reports") Then
					$iResults += 1
					$bReturn = True
				EndIf

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						ReDim $avFolders[$iFolders + 1][2]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][2]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, $iResults, $bReturn)
EndFunc   ;==>_LOBase_ReportFolderExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFolderRename
; Description ...: Rename a Report Folder.
; Syntax ........: _LOBase_ReportFolderRename(ByRef $oDoc, $sFolder, $sNewName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sFolder             - a string value. The Folder to rename, including the Sub-Folder path, if applicable. See Remarks.
;                  $sNewName            - a string value. The New name to rename the Report Folder to.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 3 Return 0 = $sNewName not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder name called in $sFolder not found in Folder or is not a Folder.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sNewName already exists in Folder.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to rename folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully renamed the Folder
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To rename a Folder inside a folder, the original Folder name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to rename FolderXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3/FolderXYZ.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFolderRename(ByRef $oDoc, $sFolder, $sNewName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oSource.hasByHierarchicalName($sFolder) Or Not $oSource.getByHierarchicalName($sFolder).supportsService("com.sun.star.sdb.Reports") Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If $oSource.hasByHierarchicalName(StringLeft($sFolder, StringInStr($sFolder, "/", 0, -1)) & $sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSource.getByHierarchicalName($sFolder).rename($sNewName)

	If Not $oSource.hasByHierarchicalName(StringLeft($sFolder, StringInStr($sFolder, "/", 0, -1)) & $sNewName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportFolderRename

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFoldersGetCount
; Description ...: Retrieve a count of Report Folders contained in the Document.
; Syntax ........: _LOBase_ReportFoldersGetCount(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, retrieves a count of all folders, including those in sub-folders.
;                  $sFolder             - [optional] a string value. Default is "". The Folder to return the count of folders for. See remarks.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder called in $sFolder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Report Folders contained in the Document as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only the count of main level Folders (not located in folders), or if $bExhaustive is called with True, it will return a count of all Folders contained in the document.
;                  You can narrow the Folder count down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get a count of Folders contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
; Related .......: _LOBase_ReportFoldersGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFoldersGetCount(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iResults = 0
	Local $asNames[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFolder <> "") Then
		$asSplit = StringSplit($sFolder, "/")

		For $i = 1 To $asSplit[0]
			If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oSource = $oSource.getByName($asSplit[$i])
			If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Next
	EndIf

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			$iResults += 1
			ReDim $avFolders[1][2]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						$iResults += 1
						ReDim $avFolders[$iFolders + 1][2]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][2]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, $iResults)
EndFunc   ;==>_LOBase_ReportFoldersGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFoldersGetNames
; Description ...: Retrieve an array of Folder Names contained in a Document.
; Syntax ........: _LOBase_ReportFoldersGetNames(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, the search looks inside sub-folders.
;                  $sFolder             - [optional] a string value. Default is "". The Sub-Folder to return the array of Folder names from. See remarks.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder called in $sFolder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Folder names contained in this Document. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only an array of main level Folder names (not located in sub-folders), or if $bExhaustive is called with True, it will return an array of all folders contained in the document.
;                  You can narrow the Folder name return down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get an array of Folders contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
;                  All Folders located in sub-folders will have the folder path prefixed to the Folder name, separated by forward slashes (/). e.g. Folder1/Folder2/Folder3.
;                  Calling $bExhaustive with True when searching inside a Folder, will get all Folder names from inside that folder, and all sub-folders.
;                  The order of the Folder names inside the folders may not necessarily be in proper order, i.e. if there are two sub folders, and folders inside the first sub-folder, the two folders will be listed first, then the folders inside the first sub-folder.
; Related .......: _LOBase_ReportFoldersGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFoldersGetNames(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iResults = 0
	Local $asNames[0], $asFolders[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFolder <> "") Then
		$asSplit = StringSplit($sFolder, "/")

		For $i = 1 To $asSplit[0]
			If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oSource = $oSource.getByName($asSplit[$i])
			If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Next
		$sFolder &= "/"
	EndIf

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			If (UBound($asFolders) >= $iResults) Then ReDim $asFolders[$iResults + 1]
			$asFolders[$iResults] = $sFolder & $asNames[$i]
			$iResults += 1
			ReDim $avFolders[1][3]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
			$avFolders[0][$iPrefix] = $asNames[$i] & "/"
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						If (UBound($asFolders) >= $iResults) Then ReDim $asFolders[$iResults + 1]
						$asFolders[$iResults] = $sFolder & $avFolders[$iCount][$iPrefix] & $asFolderList[$k]
						$iResults += 1
						ReDim $avFolders[$iFolders + 1][3]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj
						$avFolders[$iFolders][$iPrefix] = $avFolders[$iCount][$iPrefix] & $asFolderList[$k] & "/"

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][3]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($asFolders), $asFolders)
EndFunc   ;==>_LOBase_ReportFoldersGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportFooter
; Description ...: Set or Retrieve a Report's Report Footer properties.
; Syntax ........: _LOBase_ReportFooter(ByRef $oReportDoc[, $bEnabled = Null[, $sName = Null[, $iForceNewPage = Null[, $bKeepTogether = Null[, $bVisible = Null[, $iHeight = Null[, $sCondPrint = Null[, $iBackColor = Null]]]]]]]])
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the Report Footer is enabled.
;                  $sName               - [optional] a string value. Default is Null. The name of the Section.
;                  $iForceNewPage       - [optional] an integer value (0-3). Default is Null. If and when to force a new page. See Constants, $LOB_REP_FORCE_PAGE_* as defined in LibreOfficeBase_Constants.au3.
;                  $bKeepTogether       - [optional] a boolean value. Default is Null. If True, the section should be printed on one page.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the section is visible in the Report.
;                  $iHeight             - [optional] an integer value (1753-??). Default is Null. The height of the Section, in Hundredths of a Millimeter (HMM). See remarks.
;                  $sCondPrint          - [optional] a string value. Default is Null. The Conditional Print Statement.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF to set Background color to default / Background Transparent = True.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $sName not a String.
;                  @Error 1 @Extended 5 Return 0 = $iForceNewPage not an Integer, less than 0 or greater than 3. See Constants, $LOB_REP_FORCE_PAGE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bKeepTogether not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iHeight not an Integer, or less than 1753.
;                  @Error 1 @Extended 9 Return 0 = $sCondPrint not a String.
;                  @Error 1 @Extended 10 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bEnabled
;                  |                               2 = Error setting $sName
;                  |                               4 = Error setting $iForceNewPage
;                  |                               8 = Error setting $bKeepTogether
;                  |                               16 = Error setting $bVisible
;                  |                               32 = Error setting $iHeight
;                  |                               64 = Error setting $sCondPrint
;                  |                               128 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Report Footer must be enabled (turned on), before you can set or retrieve any other properties. When retrieving the current properties when the Footer is disabled, the return values will be Null, except for the Boolean value of $bEnabled.
;                  The minimum height of a Section is 1753 Hundredths of a Millimeter (HMM), the maximum is unknown, but I found that setting a large value tends to cause a freeze up/crash of the Report.
;                  Background Transparent is set automatically based on the value set for Background color. Set Background color to $LO_COLOR_OFF to set Background Transparent to True.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_ReportPageFooter, _LOBase_ReportPageHeader, _LOBase_ReportHeader, _LOBase_ReportDetail, _LOBase_ReportGroupFooter, _LOBase_ReportGroupHeader, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportFooter(ByRef $oReportDoc, $bEnabled = Null, $sName = Null, $iForceNewPage = Null, $bKeepTogether = Null, $bVisible = Null, $iHeight = Null, $sCondPrint = Null, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avProps[8]
	Local $iError = 0

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bEnabled, $sName, $iForceNewPage, $bKeepTogether, $bVisible, $iHeight, $sCondPrint, $iBackColor) Then
		If $oReportDoc.ReportFooterOn() Then
			__LO_ArrayFill($avProps, $oReportDoc.ReportFooterOn(), $oReportDoc.ReportFooter.Name(), $oReportDoc.ReportFooter.ForceNewPage(), _
					$oReportDoc.ReportFooter.KeepTogether(), $oReportDoc.ReportFooter.Visible(), $oReportDoc.ReportFooter.Height(), _
					$oReportDoc.ReportFooter.ConditionalPrintExpression(), $oReportDoc.ReportFooter.BackColor())

		Else ; Page Footer is off.
			__LO_ArrayFill($avProps, $oReportDoc.ReportFooterOn(), Null, Null, Null, Null, Null, Null, Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avProps)
	EndIf

	If ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oReportDoc.ReportFooterOn = $bEnabled
		$iError = ($oReportDoc.ReportFooterOn() = $bEnabled) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sName <> Null) Then
		If $oReportDoc.ReportFooterOn() Then
			If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oReportDoc.ReportFooter.Name = $sName
			$iError = ($oReportDoc.ReportFooter.Name() = $sName) ? ($iError) : (BitOR($iError, 2))

		Else
			$iError = BitOR($iError, 2) ; Can't set ReportFooter Values if Footer is off.
		EndIf
	EndIf

	If ($iForceNewPage <> Null) Then
		If $oReportDoc.ReportFooterOn() Then
			If Not __LO_IntIsBetween($iForceNewPage, $LOB_REP_FORCE_PAGE_NONE, $LOB_REP_FORCE_PAGE_BEFORE_AFTER_SECTION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oReportDoc.ReportFooter.ForceNewPage = $iForceNewPage
			$iError = ($oReportDoc.ReportFooter.ForceNewPage() = $iForceNewPage) ? ($iError) : (BitOR($iError, 4))

		Else
			$iError = BitOR($iError, 4) ; Can't set ReportFooter Values if Footer is off.
		EndIf
	EndIf

	If ($bKeepTogether <> Null) Then
		If $oReportDoc.ReportFooterOn() Then
			If Not IsBool($bKeepTogether) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oReportDoc.ReportFooter.KeepTogether = $bKeepTogether
			$iError = ($oReportDoc.ReportFooter.KeepTogether() = $bKeepTogether) ? ($iError) : (BitOR($iError, 8))

		Else
			$iError = BitOR($iError, 8) ; Can't set ReportFooter Values if Footer is off.
		EndIf
	EndIf

	If ($bVisible <> Null) Then
		If $oReportDoc.ReportFooterOn() Then
			If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oReportDoc.ReportFooter.Visible = $bVisible
			$iError = ($oReportDoc.ReportFooter.Visible() = $bVisible) ? ($iError) : (BitOR($iError, 16))

		Else
			$iError = BitOR($iError, 16) ; Can't set ReportFooter Values if Footer is off.
		EndIf
	EndIf

	If ($iHeight <> Null) Then
		If $oReportDoc.ReportFooterOn() Then
			If Not __LO_IntIsBetween($iHeight, 1753) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oReportDoc.ReportFooter.Height = $iHeight
			$iError = (__LO_IntIsBetween($oReportDoc.ReportFooter.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 32))

		Else
			$iError = BitOR($iError, 32) ; Can't set ReportFooter Values if Footer is off.
		EndIf
	EndIf

	If ($sCondPrint <> Null) Then
		If $oReportDoc.ReportFooterOn() Then
			If Not IsString($sCondPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$oReportDoc.ReportFooter.ConditionalPrintExpression = $sCondPrint
			$iError = ($oReportDoc.ReportFooter.ConditionalPrintExpression() = $sCondPrint) ? ($iError) : (BitOR($iError, 64))

		Else
			$iError = BitOR($iError, 64) ; Can't set ReportFooter Values if Footer is off.
		EndIf
	EndIf

	If ($iBackColor <> Null) Then
		If $oReportDoc.ReportFooterOn() Then
			If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			$oReportDoc.ReportFooter.BackColor = $iBackColor
			$iError = ($oReportDoc.ReportFooter.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 128))

		Else
			$iError = BitOR($iError, 128) ; Can't set ReportFooter Values if Footer is off.
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportGeneral
; Description ...: Set or Retrieve General Report Document properties.
; Syntax ........: _LOBase_ReportGeneral(ByRef $oReportDoc[, $sName = Null[, $iPageHeader = Null[, $iPageFooter = Null[, $bAutoGrow = Null[, $bPrintRep = Null]]]]])
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $sName               - [optional] a string value. Default is Null. The name of the Report. This is separate from the save name of the Report contained in the Database.
;                  $iPageHeader         - [optional] an integer value (0-3). Default is Null. Determines if a Page Header is printed on a page that also contains a Report Header. See Constants, $LOB_REP_PAGE_PRINT_OPT_* as defined in LibreOfficeBase_Constants.au3.
;                  $iPageFooter         - [optional] an integer value (0-3). Default is Null. Determines if a Page Footer is printed on a page that also contains a Report Footer. See Constants, $LOB_REP_PAGE_PRINT_OPT_* as defined in LibreOfficeBase_Constants.au3.
;                  $bAutoGrow           - [optional] a boolean value. Default is Null. If True, the Report will automatically grow to fit content.
;                  $bPrintRep           - [optional] a boolean value. Default is Null. If True, repeated values will be printed.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $iPageHeader not an Integer, less than 0 or greater than 3. See Constants, $LOB_REP_PAGE_PRINT_OPT_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iPageFooter not an Integer, less than 0 or greater than 3. See Constants, $LOB_REP_PAGE_PRINT_OPT_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bAutoGrow not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bPrintRep not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $iPageHeader
;                  |                               4 = Error setting $iPageFooter
;                  |                               8 = Error setting $bAutoGrow
;                  |                               16 = Error setting $bPrintRep
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_ReportData
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportGeneral(ByRef $oReportDoc, $sName = Null, $iPageHeader = Null, $iPageFooter = Null, $bAutoGrow = Null, $bPrintRep = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avReport[5]

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sName, $iPageHeader, $iPageFooter, $bAutoGrow, $bPrintRep) Then
		__LO_ArrayFill($avReport, $oReportDoc.Name(), $oReportDoc.PageHeaderOption(), $oReportDoc.PageFooterOption(), $oReportDoc.AutoGrow(), $oReportDoc.PrintRepeatedValues())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avReport)
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oReportDoc.Name = $sName
		$iError = ($oReportDoc.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iPageHeader <> Null) Then
		If Not __LO_IntIsBetween($iPageHeader, $LOB_REP_PAGE_PRINT_OPT_ALL_PAGES, $LOB_REP_PAGE_PRINT_OPT_NOT_WITH_REP_HEADER_FOOTER) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oReportDoc.PageHeaderOption = $iPageHeader
		$iError = ($oReportDoc.PageHeaderOption() = $iPageHeader) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iPageFooter <> Null) Then
		If Not __LO_IntIsBetween($iPageFooter, $LOB_REP_PAGE_PRINT_OPT_ALL_PAGES, $LOB_REP_PAGE_PRINT_OPT_NOT_WITH_REP_HEADER_FOOTER) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oReportDoc.PageFooterOption = $iPageFooter
		$iError = ($oReportDoc.PageFooterOption() = $iPageFooter) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bAutoGrow <> Null) Then
		If Not IsBool($bAutoGrow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oReportDoc.AutoGrow = $bAutoGrow
		$iError = ($oReportDoc.AutoGrow() = $bAutoGrow) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bPrintRep <> Null) Then
		If Not IsBool($bPrintRep) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oReportDoc.PrintRepeatedValues = $bPrintRep
		$iError = ($oReportDoc.PrintRepeatedValues() = $bPrintRep) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportGeneral

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportGroupAdd
; Description ...: Add a Group to the Report.
; Syntax ........: _LOBase_ReportGroupAdd(ByRef $oReportDoc[, $iPosition = Null])
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $iPosition           - [optional] an integer value. Default is Null. The position to insert the new Group. 0 Based, call Null to insert at the end.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $iPosition not an Integer, less than 0 or greater than number of Groups contained in the Report plus one.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create new Group object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed retrieve new Group Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning new Group Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportGroupAdd(ByRef $oReportDoc, $iPosition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oGroup, $oReportGroup

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iPosition) Then $iPosition = $oReportDoc.Groups.Count()

	If Not __LO_IntIsBetween($iPosition, 0, $oReportDoc.Groups.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oGroup = $oReportDoc.Groups.createGroup()
	If Not IsObj($oGroup) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oReportDoc.Groups.insertByIndex($iPosition, $oGroup)

	$oReportGroup = $oReportDoc.Groups.getByIndex($iPosition)
	If Not IsObj($oReportGroup) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oReportGroup.HeaderOn = True ; Turn Header on so it is visible to the user.

	Return SetError($__LO_STATUS_SUCCESS, 0, $oReportGroup)
EndFunc   ;==>_LOBase_ReportGroupAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportGroupDeleteByIndex
; Description ...: Delete a Group by position.
; Syntax ........: _LOBase_ReportGroupDeleteByIndex(ByRef $oReportDoc, $iGroup)
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $iGroup              - an integer value. The Index position of the Group to Delete. 0 based.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $iGroup not an Integer, less than 0 or greater than number of Groups contained in the Report.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed retrieve a count of Groups.
;                  @Error 3 @Extended 2 Return 0 = Failed to delete Group.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Returning requested Group Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_ReportGroupDeleteByObj, _LOBase_ReportGroupAdd
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportGroupDeleteByIndex(ByRef $oReportDoc, $iGroup)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iGroup, 0, $oReportDoc.Groups.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iCount = $oReportDoc.Groups.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oReportDoc.Groups.removeByIndex($iGroup)

	If (($iCount - 1) <> $oReportDoc.Groups.Count()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportGroupDeleteByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportGroupDeleteByObj
; Description ...: Delete a Group by its Object.
; Syntax ........: _LOBase_ReportGroupDeleteByObj(ByRef $oGroup)
; Parameters ....: $oGroup              - [in/out] an object. A Group object returned by a previous _LOBase_ReportGroupAdd, or _LOBase_ReportGroupGetByIndex function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oGroup not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oGroup not a Group Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed retrieve a Group Parent Object.
;                  @Error 3 @Extended 2 Return 0 = Failed retrieve a count of Groups.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete Group.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Returning requested Group Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_ReportGroupDeleteByIndex, _LOBase_ReportGroupAdd
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportGroupDeleteByObj(ByRef $oGroup)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount
	Local $oParent

	If Not IsObj($oGroup) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oGroup.supportsService("com.sun.star.report.Group") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oParent = $oGroup.Parent()
	If Not IsObj($oParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oParent.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To $iCount - 1
		If ($oParent.getByIndex($i) = $oGroup) Then
			$oParent.removeByIndex($i)
			ExitLoop
		EndIf

		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If (($iCount - 1) <> $oParent.Count()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportGroupDeleteByObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportGroupFooter
; Description ...: Set or Retrieve Group Footer settings.
; Syntax ........: _LOBase_ReportGroupFooter(ByRef $oGroup[, $bFooterOn = Null[, $sName = Null[, $iForceNewPage = Null[, $bKeepTogether = Null[, $bRepeatSec = Null[, $bVisible = Null[, $iHeight = Null[, $sCondPrint = Null[, $iBackColor = Null]]]]]]]]])
; Parameters ....: $oGroup              - [in/out] an object. A Group object returned by a previous _LOBase_ReportGroupAdd, or _LOBase_ReportGroupGetByIndex function.
;                  $bFooterOn           - [optional] a boolean value. Default is Null. If True, the Footer is enabled (on).
;                  $sName               - [optional] a string value. Default is Null. The name of the Group Footer.
;                  $iForceNewPage       - [optional] an integer value (0-3). Default is Null. If and when to force a new page. See Constants, $LOB_REP_FORCE_PAGE_* as defined in LibreOfficeBase_Constants.au3.
;                  $bKeepTogether       - [optional] a boolean value. Default is Null. If True, the section should be printed on one page.
;                  $bRepeatSec          - [optional] a boolean value. Default is Null. If True, the Group Footer section will be repeated on the next page if the section spans more than one page.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the section is visible in the Report.
;                  $iHeight             - [optional] an integer value (1753-??). Default is Null. The height of the Section, in Hundredths of a Millimeter (HMM). See remarks.
;                  $sCondPrint          - [optional] a string value. Default is Null. The Conditional Print Statement.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF to set Background color to default / Background Transparent = True.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oGroup not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oGroup not a Group Object.
;                  @Error 1 @Extended 3 Return 0 = $bFooterOn not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $sName not a String.
;                  @Error 1 @Extended 5 Return 0 = $iForceNewPage not an Integer, less than 0 or greater than 3. See Constants, $LOB_REP_FORCE_PAGE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bKeepTogether not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bRepeatSec not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iHeight not an Integer, or less than 1753.
;                  @Error 1 @Extended 10 Return 0 = $sCondPrint not a String.
;                  @Error 1 @Extended 11 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFooterOn
;                  |                               2 = Error setting $sName
;                  |                               4 = Error setting $iForceNewPage
;                  |                               8 = Error setting $bKeepTogether
;                  |                               16 = Error setting $bRepeatSec
;                  |                               32 = Error setting $bVisible
;                  |                               64 = Error setting $iHeight
;                  |                               128 = Error setting $sCondPrint
;                  |                               256 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 9 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Group Footer must be enabled (turned on), before you can set or retrieve any other properties. When retrieving the current properties when the Footer is disabled, the return values will be Null, except for the Boolean value of $bFooterOn.
;                  The minimum height of a Section is 1753 Hundredths of a Millimeter (HMM), the maximum is unknown, but I found that setting a large value tends to cause a freeze up/crash of the Report.
;                  Background Transparent is set automatically based on the value set for Background color. Set Background color to $LO_COLOR_OFF to set Background Transparent to True.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_ReportPageFooter, _LOBase_ReportPageHeader, _LOBase_ReportFooter, _LOBase_ReportHeader, _LOBase_ReportDetail, _LOBase_ReportGroupHeader, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportGroupFooter(ByRef $oGroup, $bFooterOn = Null, $sName = Null, $iForceNewPage = Null, $bKeepTogether = Null, $bRepeatSec = Null, $bVisible = Null, $iHeight = Null, $sCondPrint = Null, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avProps[9]
	Local $iError = 0

	If Not IsObj($oGroup) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oGroup.supportsService("com.sun.star.report.Group") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bFooterOn, $sName, $iForceNewPage, $bKeepTogether, $bRepeatSec, $bVisible, $iHeight, $sCondPrint, $iBackColor) Then
		If $oGroup.FooterOn() Then
			__LO_ArrayFill($avProps, $oGroup.FooterOn(), $oGroup.Footer.Name(), $oGroup.Footer.ForceNewPage(), $oGroup.Footer.KeepTogether(), _
					$oGroup.Footer.RepeatSection(), $oGroup.Footer.Visible(), $oGroup.Footer.Height(), $oGroup.Footer.ConditionalPrintExpression(), _
					$oGroup.Footer.BackColor())

		Else ; Page Footer is off.
			__LO_ArrayFill($avProps, $oGroup.FooterOn(), Null, Null, Null, Null, Null, Null, Null, Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avProps)
	EndIf

	If ($bFooterOn <> Null) Then
		If Not IsBool($bFooterOn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oGroup.FooterOn = $bFooterOn
		$iError = ($oGroup.FooterOn() = $bFooterOn) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sName <> Null) Then
		If $oGroup.FooterOn() Then
			If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oGroup.Footer.Name = $sName
			$iError = ($oGroup.Footer.Name() = $sName) ? ($iError) : (BitOR($iError, 2))

		Else
			$iError = BitOR($iError, 2) ; Can't set Footer Values if Footer is off.
		EndIf
	EndIf

	If ($iForceNewPage <> Null) Then
		If $oGroup.FooterOn() Then
			If Not __LO_IntIsBetween($iForceNewPage, $LOB_REP_FORCE_PAGE_NONE, $LOB_REP_FORCE_PAGE_BEFORE_AFTER_SECTION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oGroup.Footer.ForceNewPage = $iForceNewPage
			$iError = ($oGroup.Footer.ForceNewPage() = $iForceNewPage) ? ($iError) : (BitOR($iError, 4))

		Else
			$iError = BitOR($iError, 4) ; Can't set Footer Values if Footer is off.
		EndIf
	EndIf

	If ($bKeepTogether <> Null) Then
		If $oGroup.FooterOn() Then
			If Not IsBool($bKeepTogether) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oGroup.Footer.KeepTogether = $bKeepTogether
			$iError = ($oGroup.Footer.KeepTogether() = $bKeepTogether) ? ($iError) : (BitOR($iError, 8))

		Else
			$iError = BitOR($iError, 8) ; Can't set Footer Values if Footer is off.
		EndIf
	EndIf

	If ($bRepeatSec <> Null) Then
		If $oGroup.FooterOn() Then
			If Not IsBool($bRepeatSec) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oGroup.Footer.RepeatSection = $bRepeatSec
			$iError = ($oGroup.Footer.RepeatSection() = $bRepeatSec) ? ($iError) : (BitOR($iError, 16))

		Else
			$iError = BitOR($iError, 16) ; Can't set Footer Values if Footer is off.
		EndIf
	EndIf

	If ($bVisible <> Null) Then
		If $oGroup.FooterOn() Then
			If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oGroup.Footer.Visible = $bVisible
			$iError = ($oGroup.Footer.Visible() = $bVisible) ? ($iError) : (BitOR($iError, 32))

		Else
			$iError = BitOR($iError, 32) ; Can't set Footer Values if Footer is off.
		EndIf
	EndIf

	If ($iHeight <> Null) Then
		If $oGroup.FooterOn() Then
			If Not __LO_IntIsBetween($iHeight, 1753) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$oGroup.Footer.Height = $iHeight
			$iError = ($oGroup.Footer.Height() = $iHeight) ? ($iError) : (BitOR($iError, 64))

		Else
			$iError = BitOR($iError, 64) ; Can't set Footer Values if Footer is off.
		EndIf
	EndIf

	If ($sCondPrint <> Null) Then
		If $oGroup.FooterOn() Then
			If Not IsString($sCondPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			$oGroup.Footer.ConditionalPrintExpression = $sCondPrint
			$iError = ($oGroup.Footer.ConditionalPrintExpression() = $sCondPrint) ? ($iError) : (BitOR($iError, 128))

		Else
			$iError = BitOR($iError, 128) ; Can't set Footer Values if Footer is off.
		EndIf
	EndIf

	If ($iBackColor <> Null) Then
		If $oGroup.FooterOn() Then
			If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

			$oGroup.Footer.BackColor = $iBackColor
			$iError = ($oGroup.Footer.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 256))

		Else
			$iError = BitOR($iError, 256) ; Can't set Footer Values if Footer is off.
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportGroupFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportGroupGetByIndex
; Description ...: Retrieve a Group Object by position.
; Syntax ........: _LOBase_ReportGroupGetByIndex(ByRef $oReportDoc, $iReport)
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $iReport             - an integer value. The index position for the Group to retrieve the Object for. 0 Based.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $iReport not an Integer, less than 0 or greater than number of Groups contained in the Report.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed retrieve Group Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Group Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportGroupGetByIndex(ByRef $oReportDoc, $iReport)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oGroup

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iReport, 0, $oReportDoc.Groups.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oGroup = $oReportDoc.Groups.getByIndex($iReport)
	If Not IsObj($oGroup) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oGroup)
EndFunc   ;==>_LOBase_ReportGroupGetByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportGroupHeader
; Description ...: Set or Retrieve Group Header settings.
; Syntax ........: _LOBase_ReportGroupHeader(ByRef $oGroup[, $bHeaderOn = Null[, $sName = Null[, $iForceNewPage = Null[, $bKeepTogether = Null[, $bRepeatSec = Null[, $bVisible = Null[, $iHeight = Null[, $sCondPrint = Null[, $iBackColor = Null]]]]]]]]])
; Parameters ....: $oGroup              - [in/out] an object. A Group object returned by a previous _LOBase_ReportGroupAdd, or _LOBase_ReportGroupGetByIndex function.
;                  $bHeaderOn           - [optional] a boolean value. Default is Null. If True, the Header is enabled (on).
;                  $sName               - [optional] a string value. Default is Null. The name of the Group Header.
;                  $iForceNewPage       - [optional] an integer value (0-3). Default is Null. If and when to force a new page. See Constants, $LOB_REP_FORCE_PAGE_* as defined in LibreOfficeBase_Constants.au3.
;                  $bKeepTogether       - [optional] a boolean value. Default is Null. If True, the section should be printed on one page.
;                  $bRepeatSec          - [optional] a boolean value. Default is Null. If True, the Group Header section will be repeated on the next page if the section spans more than one page.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the section is visible in the Report.
;                  $iHeight             - [optional] an integer value (1753-??). Default is Null. The height of the Section, in Hundredths of a Millimeter (HMM). See remarks.
;                  $sCondPrint          - [optional] a string value. Default is Null. The Conditional Print Statement.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF to set Background color to default / Background Transparent = True.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oGroup not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oGroup not a Group Object.
;                  @Error 1 @Extended 3 Return 0 = $bHeaderOn not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $sName not a String.
;                  @Error 1 @Extended 5 Return 0 = $iForceNewPage not an Integer, less than 0 or greater than 3. See Constants, $LOB_REP_FORCE_PAGE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bKeepTogether not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bRepeatSec not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iHeight not an Integer, or less than 1753.
;                  @Error 1 @Extended 10 Return 0 = $sCondPrint not a String.
;                  @Error 1 @Extended 11 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bHeaderOn
;                  |                               2 = Error setting $sName
;                  |                               4 = Error setting $iForceNewPage
;                  |                               8 = Error setting $bKeepTogether
;                  |                               16 = Error setting $bRepeatSec
;                  |                               32 = Error setting $bVisible
;                  |                               64 = Error setting $iHeight
;                  |                               128 = Error setting $sCondPrint
;                  |                               256 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 9 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Group Header must be enabled (turned on), before you can set or retrieve any other properties. When retrieving the current properties when the Header is disabled, the return values will be Null, except for the Boolean value of $bHeaderOn.
;                  The minimum height of a Section is 1753 Hundredths of a Millimeter (HMM), the maximum is unknown, but I found that setting a large value tends to cause a freeze up/crash of the Report.
;                  Background Transparent is set automatically based on the value set for Background color. Set Background color to $LO_COLOR_OFF to set Background Transparent to True.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_ReportPageFooter, _LOBase_ReportPageHeader, _LOBase_ReportFooter, _LOBase_ReportHeader, _LOBase_ReportDetail, _LOBase_ReportGroupFooter, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportGroupHeader(ByRef $oGroup, $bHeaderOn = Null, $sName = Null, $iForceNewPage = Null, $bKeepTogether = Null, $bRepeatSec = Null, $bVisible = Null, $iHeight = Null, $sCondPrint = Null, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avProps[9]
	Local $iError = 0

	If Not IsObj($oGroup) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oGroup.supportsService("com.sun.star.report.Group") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bHeaderOn, $sName, $iForceNewPage, $bKeepTogether, $bRepeatSec, $bVisible, $iHeight, $sCondPrint, $iBackColor) Then
		If $oGroup.HeaderOn() Then
			__LO_ArrayFill($avProps, $oGroup.HeaderOn(), $oGroup.Header.Name(), $oGroup.Header.ForceNewPage(), $oGroup.Header.KeepTogether(), _
					$oGroup.Header.RepeatSection(), $oGroup.Header.Visible(), $oGroup.Header.Height(), $oGroup.Header.ConditionalPrintExpression(), _
					$oGroup.Header.BackColor())

		Else ; Page Header is off.
			__LO_ArrayFill($avProps, $oGroup.HeaderOn(), Null, Null, Null, Null, Null, Null, Null, Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avProps)
	EndIf

	If ($bHeaderOn <> Null) Then
		If Not IsBool($bHeaderOn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oGroup.HeaderOn = $bHeaderOn
		$iError = ($oGroup.HeaderOn() = $bHeaderOn) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sName <> Null) Then
		If $oGroup.HeaderOn() Then
			If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oGroup.Header.Name = $sName
			$iError = ($oGroup.Header.Name() = $sName) ? ($iError) : (BitOR($iError, 2))

		Else
			$iError = BitOR($iError, 2) ; Can't set Header Values if Header is off.
		EndIf
	EndIf

	If ($iForceNewPage <> Null) Then
		If $oGroup.HeaderOn() Then
			If Not __LO_IntIsBetween($iForceNewPage, $LOB_REP_FORCE_PAGE_NONE, $LOB_REP_FORCE_PAGE_BEFORE_AFTER_SECTION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oGroup.Header.ForceNewPage = $iForceNewPage
			$iError = ($oGroup.Header.ForceNewPage() = $iForceNewPage) ? ($iError) : (BitOR($iError, 4))

		Else
			$iError = BitOR($iError, 4) ; Can't set Header Values if Header is off.
		EndIf
	EndIf

	If ($bKeepTogether <> Null) Then
		If $oGroup.HeaderOn() Then
			If Not IsBool($bKeepTogether) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oGroup.Header.KeepTogether = $bKeepTogether
			$iError = ($oGroup.Header.KeepTogether() = $bKeepTogether) ? ($iError) : (BitOR($iError, 8))

		Else
			$iError = BitOR($iError, 8) ; Can't set Header Values if Header is off.
		EndIf
	EndIf

	If ($bRepeatSec <> Null) Then
		If $oGroup.HeaderOn() Then
			If Not IsBool($bRepeatSec) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oGroup.Header.RepeatSection = $bRepeatSec
			$iError = ($oGroup.Header.RepeatSection() = $bRepeatSec) ? ($iError) : (BitOR($iError, 16))

		Else
			$iError = BitOR($iError, 16) ; Can't set Header Values if Header is off.
		EndIf
	EndIf

	If ($bVisible <> Null) Then
		If $oGroup.HeaderOn() Then
			If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oGroup.Header.Visible = $bVisible
			$iError = ($oGroup.Header.Visible() = $bVisible) ? ($iError) : (BitOR($iError, 32))

		Else
			$iError = BitOR($iError, 32) ; Can't set Header Values if Header is off.
		EndIf
	EndIf

	If ($iHeight <> Null) Then
		If $oGroup.HeaderOn() Then
			If Not __LO_IntIsBetween($iHeight, 1753) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$oGroup.Header.Height = $iHeight
			$iError = ($oGroup.Header.Height() = $iHeight) ? ($iError) : (BitOR($iError, 64))

		Else
			$iError = BitOR($iError, 64) ; Can't set Header Values if Header is off.
		EndIf
	EndIf

	If ($sCondPrint <> Null) Then
		If $oGroup.HeaderOn() Then
			If Not IsString($sCondPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			$oGroup.Header.ConditionalPrintExpression = $sCondPrint
			$iError = ($oGroup.Header.ConditionalPrintExpression() = $sCondPrint) ? ($iError) : (BitOR($iError, 128))

		Else
			$iError = BitOR($iError, 128) ; Can't set Header Values if Header is off.
		EndIf
	EndIf

	If ($iBackColor <> Null) Then
		If $oGroup.HeaderOn() Then
			If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

			$oGroup.Header.BackColor = $iBackColor
			$iError = ($oGroup.Header.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 256))

		Else
			$iError = BitOR($iError, 256) ; Can't set Header Values if Header is off.
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportGroupHeader

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportGroupPosition
; Description ...: Set or Retrieve the Group's position in the list of Groups.
; Syntax ........: _LOBase_ReportGroupPosition(ByRef $oGroup[, $iPos = Null])
; Parameters ....: $oGroup              - [in/out] an object. A Group object returned by a previous _LOBase_ReportGroupAdd, or _LOBase_ReportGroupGetByIndex function.
;                  $iPos                - [optional] an integer value. Default is Null. The position of the in the list of Groups. 0 Based. See Remarks.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oGroup not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oGroup not a Group Object.
;                  @Error 1 @Extended 3 Return 0 = $iPos not an Integer, less than 0 or greater than number of Groups.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Group's Parent Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve count of Groups.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify Group's current Position.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Group's new Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to delete old Group.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Group was successfully moved.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were called with Null, returning current Position as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Group will be moved to the position before that called in $iPos. Thus to move a Group to the end of the list call $iPos with the total count of Groups, i.e., index of the last Group + 1.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current Position.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportGroupPosition(ByRef $oGroup, $iPos = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount, $iCurPos
	Local $oParent, $oNewGroup

	If Not IsObj($oGroup) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oGroup.supportsService("com.sun.star.report.Group") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oParent = $oGroup.Parent()
	If Not IsObj($oParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oParent.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To $iCount - 1
		If ($oParent.getByIndex($i) = $oGroup) Then
			$iCurPos = $i
			ExitLoop
		EndIf

		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If Not IsInt($iCurPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If __LO_VarsAreNull($iPos) Then Return SetError($__LO_STATUS_SUCCESS, 0, $iCurPos)

	If Not __LO_IntIsBetween($iPos, 0, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oParent.insertByIndex($iPos, $oGroup)

	$oNewGroup = $oParent.getByIndex($iPos)
	If Not IsObj($oNewGroup) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	If ($iPos <= $iCurPos) Then $iCurPos += 1 ; If inserting before the current position, increase current position count to match new position.
	$oParent.removeByIndex($iCurPos)

	If ($iCount <> $oParent.Count()) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$oGroup = $oNewGroup

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportGroupPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportGroupsGetCount
; Description ...: Retrieve a count of Report Groups.
; Syntax ........: _LOBase_ReportGroupsGetCount(ByRef $oReportDoc)
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve count of Groups.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning total number of Groups contained in the Report.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportGroupsGetCount(ByRef $oReportDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iCount = $oReportDoc.Groups.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOBase_ReportGroupsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportGroupSort
; Description ...: Set or Retrieve a Group's Sorting settings.
; Syntax ........: _LOBase_ReportGroupSort(ByRef $oGroup[, $sField = Null[, $bSortAsc = Null[, $iGroupOn = Null[, $iGroupInt = Null[, $iKeepTogether = Null]]]]])
; Parameters ....: $oGroup              - [in/out] an object. A Group object returned by a previous _LOBase_ReportGroupAdd, or _LOBase_ReportGroupGetByIndex function.
;                  $sField              - [optional] a string value. Default is Null. The Column name or Expression. See remarks.
;                  $bSortAsc            - [optional] a boolean value. Default is Null. If True, the Group is sorted in Ascending order. Else in Descending order.
;                  $iGroupOn            - [optional] an integer value (0-9). Default is Null. How to Group the Data. See Constants, $LOB_REP_GROUP_ON_* as defined in LibreOfficeBase_Constants.au3.
;                  $iGroupInt           - [optional] an integer value (0-100). Default is Null. The Group Interval value.
;                  $iKeepTogether       - [optional] an integer value (0-2). Default is Null. Whether or not, and how to keep Data together on one page. See Constants, $LOB_REP_KEEP_TOG_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oGroup not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oGroup not a Group Object.
;                  @Error 1 @Extended 3 Return 0 = $sField not a String.
;                  @Error 1 @Extended 4 Return 0 = $bSortAsc not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iGroupOn not an Integer, less than 0 or greater than 9. See Constants, $LOB_REP_GROUP_ON_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iGroupInt not an Integer, less than 0 or greater than 100.
;                  @Error 1 @Extended 7 Return 0 = $iKeepTogether not an Integer, less than 0 or greater than 2. See Constants, $LOB_REP_KEEP_TOG_* as defined in LibreOfficeBase_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sField
;                  |                               2 = Error setting $bSortAsc
;                  |                               4 = Error setting $iGroupOn
;                  |                               8 = Error setting $iGroupInt
;                  |                               16 = Error setting $iKeepTogether
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  It is the User's responsibility for the accuracy of names etc called in $sField, i.e. Column name, etc.
;                  It is the User's responsibility to use appropriate values for $iGroupOn based upon the type of field.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportGroupSort(ByRef $oGroup, $sField = Null, $bSortAsc = Null, $iGroupOn = Null, $iGroupInt = Null, $iKeepTogether = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avProps[5]
	Local $iError = 0

	If Not IsObj($oGroup) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oGroup.supportsService("com.sun.star.report.Group") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sField, $bSortAsc, $iGroupOn, $iGroupInt, $iKeepTogether) Then
		__LO_ArrayFill($avProps, $oGroup.Expression(), $oGroup.SortAscending(), $oGroup.GroupOn(), $oGroup.GroupInterval(), $oGroup.KeepTogether())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avProps)
	EndIf

	If ($sField <> Null) Then
		If Not IsString($sField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oGroup.Expression = $sField
		$iError = ($oGroup.Expression() = $sField) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bSortAsc <> Null) Then
		If Not IsBool($bSortAsc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oGroup.SortAscending = $bSortAsc
		$iError = ($oGroup.SortAscending() = $bSortAsc) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iGroupOn <> Null) Then
		If Not __LO_IntIsBetween($iGroupOn, $LOB_REP_GROUP_ON_DEFAULT, $LOB_REP_GROUP_ON_INTERVAL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oGroup.GroupOn = $iGroupOn
		$iError = ($oGroup.GroupOn() = $iGroupOn) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iGroupInt <> Null) Then
		If Not __LO_IntIsBetween($iGroupInt, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oGroup.GroupInterval = $iGroupInt
		$iError = ($oGroup.GroupInterval() = $iGroupInt) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iKeepTogether <> Null) Then
		If Not __LO_IntIsBetween($iKeepTogether, $LOB_REP_KEEP_TOG_NO, $LOB_REP_KEEP_TOG_WITH_FIRST_DETAIL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oGroup.KeepTogether = $iKeepTogether
		$iError = ($oGroup.KeepTogether() = $iKeepTogether) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportGroupSort

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportHeader
; Description ...: Set or Retrieve a Report's Report Header properties.
; Syntax ........: _LOBase_ReportHeader(ByRef $oReportDoc[, $bEnabled = Null[, $sName = Null[, $iForceNewPage = Null[, $bKeepTogether = Null[, $bVisible = Null[, $iHeight = Null[, $sCondPrint = Null[, $iBackColor = Null]]]]]]]])
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the Report Header is enabled.
;                  $sName               - [optional] a string value. Default is Null. The name of the Section.
;                  $iForceNewPage       - [optional] an integer value (0-3). Default is Null. If and when to force a new page. See Constants, $LOB_REP_FORCE_PAGE_* as defined in LibreOfficeBase_Constants.au3.
;                  $bKeepTogether       - [optional] a boolean value. Default is Null. If True, the section should be printed on one page.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the section is visible in the Report.
;                  $iHeight             - [optional] an integer value (1753-??). Default is Null. The height of the Section, in Hundredths of a Millimeter (HMM). See remarks.
;                  $sCondPrint          - [optional] a string value. Default is Null. The Conditional Print Statement.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF to set Background color to default / Background Transparent = True.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $sName not a String.
;                  @Error 1 @Extended 5 Return 0 = $iForceNewPage not an Integer, less than 0 or greater than 3. See Constants, $LOB_REP_FORCE_PAGE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bKeepTogether not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iHeight not an Integer, or less than 1753.
;                  @Error 1 @Extended 9 Return 0 = $sCondPrint not a String.
;                  @Error 1 @Extended 10 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bEnabled
;                  |                               2 = Error setting $sName
;                  |                               4 = Error setting $iForceNewPage
;                  |                               8 = Error setting $bKeepTogether
;                  |                               16 = Error setting $bVisible
;                  |                               32 = Error setting $iHeight
;                  |                               64 = Error setting $sCondPrint
;                  |                               128 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Report Header must be enabled (turned on), before you can set or retrieve any other properties. When retrieving the current properties when the Header is disabled, the return values will be Null, except for the Boolean value of $bEnabled.
;                  The minimum height of a Section is 1753 Hundredths of a Millimeter (HMM), the maximum is unknown, but I found that setting a large value tends to cause a freeze up/crash of the Report.
;                  Background Transparent is set automatically based on the value set for Background color. Set Background color to $LO_COLOR_OFF to set Background Transparent to True.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_ReportPageFooter, _LOBase_ReportPageHeader, _LOBase_ReportFooter, _LOBase_ReportDetail, _LOBase_ReportGroupFooter, _LOBase_ReportGroupHeader, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportHeader(ByRef $oReportDoc, $bEnabled = Null, $sName = Null, $iForceNewPage = Null, $bKeepTogether = Null, $bVisible = Null, $iHeight = Null, $sCondPrint = Null, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avProps[8]
	Local $iError = 0

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bEnabled, $sName, $iForceNewPage, $bKeepTogether, $bVisible, $iHeight, $sCondPrint, $iBackColor) Then
		If $oReportDoc.ReportHeaderOn() Then
			__LO_ArrayFill($avProps, $oReportDoc.ReportHeaderOn(), $oReportDoc.ReportHeader.Name(), $oReportDoc.ReportHeader.ForceNewPage(), _
					$oReportDoc.ReportHeader.KeepTogether(), $oReportDoc.ReportHeader.Visible(), $oReportDoc.ReportHeader.Height(), _
					$oReportDoc.ReportHeader.ConditionalPrintExpression(), $oReportDoc.ReportHeader.BackColor())

		Else ; Page Header is off.
			__LO_ArrayFill($avProps, $oReportDoc.ReportHeaderOn(), Null, Null, Null, Null, Null, Null, Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avProps)
	EndIf

	If ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oReportDoc.ReportHeaderOn = $bEnabled
		$iError = ($oReportDoc.ReportHeaderOn() = $bEnabled) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sName <> Null) Then
		If $oReportDoc.ReportHeaderOn() Then
			If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oReportDoc.ReportHeader.Name = $sName
			$iError = ($oReportDoc.ReportHeader.Name() = $sName) ? ($iError) : (BitOR($iError, 2))

		Else
			$iError = BitOR($iError, 2) ; Can't set ReportHeader Values if Header is off.
		EndIf
	EndIf

	If ($iForceNewPage <> Null) Then
		If $oReportDoc.ReportHeaderOn() Then
			If Not __LO_IntIsBetween($iForceNewPage, $LOB_REP_FORCE_PAGE_NONE, $LOB_REP_FORCE_PAGE_BEFORE_AFTER_SECTION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oReportDoc.ReportHeader.ForceNewPage = $iForceNewPage
			$iError = ($oReportDoc.ReportHeader.ForceNewPage() = $iForceNewPage) ? ($iError) : (BitOR($iError, 4))

		Else
			$iError = BitOR($iError, 4) ; Can't set ReportHeader Values if Header is off.
		EndIf
	EndIf

	If ($bKeepTogether <> Null) Then
		If $oReportDoc.ReportHeaderOn() Then
			If Not IsBool($bKeepTogether) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oReportDoc.ReportHeader.KeepTogether = $bKeepTogether
			$iError = ($oReportDoc.ReportHeader.KeepTogether() = $bKeepTogether) ? ($iError) : (BitOR($iError, 8))

		Else
			$iError = BitOR($iError, 8) ; Can't set ReportHeader Values if Header is off.
		EndIf
	EndIf

	If ($bVisible <> Null) Then
		If $oReportDoc.ReportHeaderOn() Then
			If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oReportDoc.ReportHeader.Visible = $bVisible
			$iError = ($oReportDoc.ReportHeader.Visible() = $bVisible) ? ($iError) : (BitOR($iError, 16))

		Else
			$iError = BitOR($iError, 16) ; Can't set ReportHeader Values if Header is off.
		EndIf
	EndIf

	If ($iHeight <> Null) Then
		If $oReportDoc.ReportHeaderOn() Then
			If Not __LO_IntIsBetween($iHeight, 1753) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oReportDoc.ReportHeader.Height = $iHeight
			$iError = (__LO_IntIsBetween($oReportDoc.ReportHeader.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 32))

		Else
			$iError = BitOR($iError, 32) ; Can't set ReportHeader Values if Header is off.
		EndIf
	EndIf

	If ($sCondPrint <> Null) Then
		If $oReportDoc.ReportHeaderOn() Then
			If Not IsString($sCondPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$oReportDoc.ReportHeader.ConditionalPrintExpression = $sCondPrint
			$iError = ($oReportDoc.ReportHeader.ConditionalPrintExpression() = $sCondPrint) ? ($iError) : (BitOR($iError, 64))

		Else
			$iError = BitOR($iError, 64) ; Can't set ReportHeader Values if Header is off.
		EndIf
	EndIf

	If ($iBackColor <> Null) Then
		If $oReportDoc.ReportHeaderOn() Then
			If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			$oReportDoc.ReportHeader.BackColor = $iBackColor
			$iError = ($oReportDoc.ReportHeader.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 128))

		Else
			$iError = BitOR($iError, 128) ; Can't set ReportHeader Values if Header is off.
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportHeader

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportIsModified
; Description ...: Test whether the Report has been modified since being created or since the last save.
; Syntax ........: _LOBase_ReportIsModified(ByRef $oReportDoc)
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the Report has been modified since last being saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_ReportSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportIsModified(ByRef $oReportDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oReportDoc.isModified())
EndFunc   ;==>_LOBase_ReportIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportOpen
; Description ...: Open a Report Document
; Syntax ........: _LOBase_ReportOpen(ByRef $oConnection, $sName[, $bDesign = True[, $bHidden = False]])
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sName               - a string value. The Report name to Open. See remarks.
;                  $bDesign             - [optional] a boolean value. Default is True. If True, the Report is opened in Design mode.
;                  $bHidden             - [optional] a boolean value. Default is False. If True, the Report document will be invisible when opened.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = $bDesign not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = Report name called in $sName not found in Folder.
;                  @Error 1 @Extended 6 Return 0 = Name called in $sName not a Report.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to open Report Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning opened Report Document's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To open a Report located inside a folder, the Report name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to open ReportXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sName with the following path: Folder1/Folder2/Folder3/ReportXYZ.
; Related .......: _LOBase_ReportClose, _LOBase_ReportConnect, _LOBase_ReportsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportOpen(ByRef $oConnection, $sName, $bDesign = True, $bHidden = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oReportDoc
	Local $aArgs[1]

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bDesign) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSource = $oConnection.Parent.DatabaseDocument.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not $oSource.Parent.CurrentController.isConnected() Then $oSource.Parent.CurrentController.connect()

	If Not $oSource.hasByHierarchicalName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not $oSource.getByHierarchicalName($sName).supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$aArgs[0] = __LO_SetPropertyValue("Hidden", $bHidden)

	$oReportDoc = $oSource.Parent.CurrentController.loadComponentWithArguments($LOB_SUB_COMP_TYPE_REPORT, $sName, $bDesign, $aArgs)
	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oReportDoc)
EndFunc   ;==>_LOBase_ReportOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportPageFooter
; Description ...: Set or Retrieve a Report's Page Footer properties.
; Syntax ........: _LOBase_ReportPageFooter(ByRef $oReportDoc[, $bEnabled = Null[, $sName = Null[, $bVisible = Null[, $iHeight = Null[, $sCondPrint = Null[, $iBackColor = Null]]]]]])
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the Page Footer is enabled.
;                  $sName               - [optional] a string value. Default is Null. The name of the Section.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the section is visible in the Report.
;                  $iHeight             - [optional] an integer value. Default is Null. (1753-??). Default is Null. The height of the Section, in Hundredths of a Millimeter (HMM). See remarks.
;                  $sCondPrint          - [optional] a string value. Default is Null. The Conditional Print Statement.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF to set Background color to default / Background Transparent = True.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $sName not a String.
;                  @Error 1 @Extended 5 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iHeight not an Integer, or less than 1753.
;                  @Error 1 @Extended 7 Return 0 = $sCondPrint not a String.
;                  @Error 1 @Extended 8 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bEnabled
;                  |                               2 = Error setting $sName
;                  |                               4 = Error setting $bVisible
;                  |                               8 = Error setting $iHeight
;                  |                               16 = Error setting $sCondPrint
;                  |                               32 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Page Header must be enabled (turned on), before you can set or retrieve any other properties. When retrieving the current properties when the Footer is disabled, the return values will be Null, except for the Boolean value of $bEnabled.
;                  The minimum height of a Section is 1753 Hundredths of a Millimeter (HMM), the maximum is unknown, but I found that setting a large value tends to cause a freeze up/crash of the Report.
;                  Background Transparent is set automatically based on the value set for Background color. Set Background color to $LO_COLOR_OFF to set Background Transparent to True.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_ReportPageHeader, _LOBase_ReportFooter, _LOBase_ReportHeader, _LOBase_ReportDetail, _LOBase_ReportGroupFooter, _LOBase_ReportGroupHeader, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportPageFooter(ByRef $oReportDoc, $bEnabled = Null, $sName = Null, $bVisible = Null, $iHeight = Null, $sCondPrint = Null, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avProps[6]
	Local $iError = 0

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bEnabled, $sName, $bVisible, $iHeight, $sCondPrint, $iBackColor) Then
		If $oReportDoc.PageFooterOn() Then
			__LO_ArrayFill($avProps, $oReportDoc.PageFooterOn(), $oReportDoc.PageFooter.Name(), $oReportDoc.PageFooter.Visible(), $oReportDoc.PageFooter.Height(), _
					$oReportDoc.PageFooter.ConditionalPrintExpression(), $oReportDoc.PageFooter.BackColor())

		Else ; Page Footer is off.
			__LO_ArrayFill($avProps, $oReportDoc.PageFooterOn(), Null, Null, Null, Null, Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avProps)
	EndIf

	If ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oReportDoc.PageFooterOn = $bEnabled
		$iError = ($oReportDoc.PageFooterOn() = $bEnabled) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sName <> Null) Then
		If $oReportDoc.PageFooterOn() Then
			If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oReportDoc.PageFooter.Name = $sName
			$iError = ($oReportDoc.PageFooter.Name() = $sName) ? ($iError) : (BitOR($iError, 2))

		Else
			$iError = BitOR($iError, 2) ; Can't set PageFooter Values if Footer is off.
		EndIf
	EndIf

	If ($bVisible <> Null) Then
		If $oReportDoc.PageFooterOn() Then
			If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oReportDoc.PageFooter.Visible = $bVisible
			$iError = ($oReportDoc.PageFooter.Visible() = $bVisible) ? ($iError) : (BitOR($iError, 4))

		Else
			$iError = BitOR($iError, 4) ; Can't set PageFooter Values if Footer is off.
		EndIf
	EndIf

	If ($iHeight <> Null) Then
		If $oReportDoc.PageFooterOn() Then
			If Not __LO_IntIsBetween($iHeight, 1753) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oReportDoc.PageFooter.Height = $iHeight
			$iError = (__LO_IntIsBetween($oReportDoc.PageFooter.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 8))

		Else
			$iError = BitOR($iError, 8) ; Can't set PageFooter Values if Footer is off.
		EndIf
	EndIf

	If ($sCondPrint <> Null) Then
		If $oReportDoc.PageFooterOn() Then
			If Not IsString($sCondPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oReportDoc.PageFooter.ConditionalPrintExpression = $sCondPrint
			$iError = ($oReportDoc.PageFooter.ConditionalPrintExpression() = $sCondPrint) ? ($iError) : (BitOR($iError, 16))

		Else
			$iError = BitOR($iError, 16) ; Can't set PageFooter Values if Footer is off.
		EndIf
	EndIf

	If ($iBackColor <> Null) Then
		If $oReportDoc.PageFooterOn() Then
			If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oReportDoc.PageFooter.BackColor = $iBackColor
			$iError = ($oReportDoc.PageFooter.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 32))

		Else
			$iError = BitOR($iError, 32) ; Can't set PageFooter Values if Footer is off.
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportPageFooter

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportPageHeader
; Description ...: Set or Retrieve a Report's Page Header properties.
; Syntax ........: _LOBase_ReportPageHeader(ByRef $oReportDoc[, $bEnabled = Null[, $sName = Null[, $bVisible = Null[, $iHeight = Null[, $sCondPrint = Null[, $iBackColor = Null]]]]]])
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $bEnabled            - [optional] a boolean value. Default is Null. If True, the Page Header is enabled.
;                  $sName               - [optional] a string value. Default is Null. The name of the Section.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the section is visible in the Report.
;                  $iHeight             - [optional] an integer value. Default is Null. (1753-??). Default is Null. The height of the Section, in Hundredths of a Millimeter (HMM). See remarks.
;                  $sCondPrint          - [optional] a string value. Default is Null. The Conditional Print Statement.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF to set Background color to default / Background Transparent = True.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $bEnabled not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $sName not a String.
;                  @Error 1 @Extended 5 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iHeight not an Integer, or less than 1753.
;                  @Error 1 @Extended 7 Return 0 = $sCondPrint not a String.
;                  @Error 1 @Extended 8 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bEnabled
;                  |                               2 = Error setting $sName
;                  |                               4 = Error setting $bVisible
;                  |                               8 = Error setting $iHeight
;                  |                               16 = Error setting $sCondPrint
;                  |                               32 = Error setting $iBackColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Page Header must be enabled (turned on), before you can set or retrieve any other properties. When retrieving the current properties when the Header is disabled, the return values will be Null, except for the Boolean value of $bEnabled.
;                  The minimum height of a Section is 1753 Hundredths of a Millimeter (HMM), the maximum is unknown, but I found that setting a large value tends to cause a freeze up/crash of the Report.
;                  Background Transparent is set automatically based on the value set for Background color. Set Background color to $LO_COLOR_OFF to set Background Transparent to True.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_ReportPageFooter, _LOBase_ReportFooter, _LOBase_ReportHeader, _LOBase_ReportDetail, _LOBase_ReportGroupFooter, _LOBase_ReportGroupHeader, _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportPageHeader(ByRef $oReportDoc, $bEnabled = Null, $sName = Null, $bVisible = Null, $iHeight = Null, $sCondPrint = Null, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avProps[6]
	Local $iError = 0

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bEnabled, $sName, $bVisible, $iHeight, $sCondPrint, $iBackColor) Then
		If $oReportDoc.PageHeaderOn() Then
			__LO_ArrayFill($avProps, $oReportDoc.PageHeaderOn(), $oReportDoc.PageHeader.Name(), $oReportDoc.PageHeader.Visible(), $oReportDoc.PageHeader.Height(), _
					$oReportDoc.PageHeader.ConditionalPrintExpression(), $oReportDoc.PageHeader.BackColor())

		Else ; Page Header is off.
			__LO_ArrayFill($avProps, $oReportDoc.PageHeaderOn(), Null, Null, Null, Null, Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avProps)
	EndIf

	If ($bEnabled <> Null) Then
		If Not IsBool($bEnabled) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oReportDoc.PageHeaderOn = $bEnabled
		$iError = ($oReportDoc.PageHeaderOn() = $bEnabled) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sName <> Null) Then
		If $oReportDoc.PageHeaderOn() Then
			If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oReportDoc.PageHeader.Name = $sName
			$iError = ($oReportDoc.PageHeader.Name() = $sName) ? ($iError) : (BitOR($iError, 2))

		Else
			$iError = BitOR($iError, 2) ; Can't set PageHeader Values if Header is off.
		EndIf
	EndIf

	If ($bVisible <> Null) Then
		If $oReportDoc.PageHeaderOn() Then
			If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oReportDoc.PageHeader.Visible = $bVisible
			$iError = ($oReportDoc.PageHeader.Visible() = $bVisible) ? ($iError) : (BitOR($iError, 4))

		Else
			$iError = BitOR($iError, 4) ; Can't set PageHeader Values if Header is off.
		EndIf
	EndIf

	If ($iHeight <> Null) Then
		If $oReportDoc.PageHeaderOn() Then
			If Not __LO_IntIsBetween($iHeight, 1753) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oReportDoc.PageHeader.Height = $iHeight
			$iError = (__LO_IntIsBetween($oReportDoc.PageHeader.Height(), $iHeight - 1, $iHeight + 1)) ? ($iError) : (BitOR($iError, 8))

		Else
			$iError = BitOR($iError, 8) ; Can't set PageHeader Values if Header is off.
		EndIf
	EndIf

	If ($sCondPrint <> Null) Then
		If $oReportDoc.PageHeaderOn() Then
			If Not IsString($sCondPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oReportDoc.PageHeader.ConditionalPrintExpression = $sCondPrint
			$iError = ($oReportDoc.PageHeader.ConditionalPrintExpression() = $sCondPrint) ? ($iError) : (BitOR($iError, 16))

		Else
			$iError = BitOR($iError, 16) ; Can't set PageHeader Values if Header is off.
		EndIf
	EndIf

	If ($iBackColor <> Null) Then
		If $oReportDoc.PageHeaderOn() Then
			If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oReportDoc.PageHeader.BackColor = $iBackColor
			$iError = ($oReportDoc.PageHeader.BackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 32))

		Else
			$iError = BitOR($iError, 32) ; Can't set PageHeader Values if Header is off.
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_ReportPageHeader

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportRename
; Description ...: Rename a Report.
; Syntax ........: _LOBase_ReportRename(ByRef $oDoc, $sReport, $sNewName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sReport             - a string value. The Report to rename, including the Sub-Folder path, if applicable. See Remarks.
;                  $sNewName            - a string value. The New name to rename the Report to.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sReport not a String.
;                  @Error 1 @Extended 3 Return 0 = $sNewName not a String.
;                  @Error 1 @Extended 4 Return 0 = Report name called in $sReport not found in Folder or is not a Report.
;                  @Error 1 @Extended 5 Return 0 = Name called in $sNewName already exists in Folder.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to rename Report.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully renamed the Report.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To rename a Report inside a folder, the original Report name MUST be prefixed by the folder path, separated by forward slashes (/). e.g. to rename ReportXYZ contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sReport with the following path: Folder1/Folder2/Folder3/ReportXYZ.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportRename(ByRef $oDoc, $sReport, $sNewName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sReport) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oSource.hasByHierarchicalName($sReport) Or Not $oSource.getByHierarchicalName($sReport).supportsService("com.sun.star.ucb.Content") Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If $oSource.hasByHierarchicalName(StringLeft($sReport, StringInStr($sReport, "/", 0, -1)) & $sNewName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oSource.getByHierarchicalName($sReport).rename($sNewName)

	If Not $oSource.hasByHierarchicalName(StringLeft($sReport, StringInStr($sReport, "/", 0, -1)) & $sNewName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportRename

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportSave
; Description ...: Save any changes made to a Document.
; Syntax ........: _LOBase_ReportSave(ByRef $oReportDoc)
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Document called in $oReportDoc is read only.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Report Document's properties.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify Report in Parent Document.
;                  @Error 3 @Extended 4 Return 0 = Document called in $oReportDoc not a Report Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Report was successfully saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: _LOBase_ReportIsModified
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportSave(ByRef $oReportDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSource, $oReport
	Local $tPropertiesPair

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oReportDoc.supportsService("com.sun.star.text.TextDocument") And $oReportDoc.isReadOnly() Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Nothing to save in a Read only Doc.

	If $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then ; Report is in Design mode.

		$oSource = $oReportDoc.Parent.ReportDocuments()
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$tPropertiesPair = $oSource.Parent.CurrentController.identifySubComponent($oReportDoc)
		If Not IsObj($tPropertiesPair) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oReport = $oSource.getByHierarchicalName($tPropertiesPair.Second())
		If Not IsObj($oReport) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Else ; Error, unknown document?

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oReport.Store()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_ReportSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportSectionGetObj
; Description ...: Retrieve a Section Object for one of the sections in a Report.
; Syntax ........: _LOBase_ReportSectionGetObj(ByRef $oReportDoc, $iSection)
; Parameters ....: $oReportDoc          - [in/out] an object. A Report Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen or _LOBase_ReportCreate function.
;                  $iSection            - an integer value (0-4). The section type to retrieve the Object for. See Constants, $LOB_REP_SECTION_TYPE_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oReportDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oReportDoc not a Report Document.
;                  @Error 1 @Extended 3 Return 0 = $iSection not an Integer, less than 0 or greater than 4. See Constants, $LOB_REP_SECTION_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Section Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Section Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportSectionGetObj(ByRef $oReportDoc, $iSection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSection

	If Not IsObj($oReportDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oReportDoc.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iSection, $LOB_REP_SECTION_TYPE_DETAIL, $LOB_REP_SECTION_TYPE_REPORT_HEADER) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	Switch $iSection
		Case $LOB_REP_SECTION_TYPE_DETAIL
			$oSection = $oReportDoc.Detail()

		Case $LOB_REP_SECTION_TYPE_PAGE_FOOTER
			$oSection = $oReportDoc.PageFooter()

		Case $LOB_REP_SECTION_TYPE_PAGE_HEADER
			$oSection = $oReportDoc.PageHeader()

		Case $LOB_REP_SECTION_TYPE_REPORT_FOOTER
			$oSection = $oReportDoc.ReportFooter()

		Case $LOB_REP_SECTION_TYPE_REPORT_HEADER
			$oSection = $oReportDoc.ReportHeader()
	EndSwitch

	If Not IsObj($oSection) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSection)
EndFunc   ;==>_LOBase_ReportSectionGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportsGetCount
; Description ...: Retrieve a count of Reports contained in the Document.
; Syntax ........: _LOBase_ReportsGetCount(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, retrieves a count of all Reports, including those in sub-folders.
;                  $sFolder             - [optional] a string value. Default is "". The Folder to return the count of Reports for. See remarks.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder called in $sFolder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Reports contained in the Document, as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only the count of main level Reports (not located in folders), or if $bExhaustive is called with True, the return will be a count of all Reports contained in the document.
;                  You can narrow the Report count down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get a count of Reports contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
; Related .......: _LOBase_ReportsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportsGetCount(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iReports = 0
	Local $asNames[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFolder <> "") Then
		$asSplit = StringSplit($sFolder, "/")

		For $i = 1 To $asSplit[0]
			If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oSource = $oSource.getByName($asSplit[$i])
			If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Next
	EndIf

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.ucb.Content") Then ; Report Doc.
			$iReports += 1

		ElseIf $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			ReDim $avFolders[1][2]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.ucb.Content") Then ; Report Doc.
						$iReports += 1

					ElseIf $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						ReDim $avFolders[$iFolders + 1][2]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][2]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, $iReports)
EndFunc   ;==>_LOBase_ReportsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ReportsGetNames
; Description ...: Retrieve an Array of Report Names contained in a Document.
; Syntax ........: _LOBase_ReportsGetNames(ByRef $oDoc[, $bExhaustive = True[, $sFolder = ""]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bExhaustive         - [optional] a boolean value. Default is True. If True, retrieves all Report names, including those in sub-folders.
;                  $sFolder             - [optional] a string value. Default is "". The Sub-Folder to return the array of Report names from. See remarks.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bExhaustive not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sFolder not a String.
;                  @Error 1 @Extended 4 Return 0 = Folder or Sub-Folder called in $sFolder not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Report Documents Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Destination Folder Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Array of Report and Folder names.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Report or Folder Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve Array of Report and Folder names for Sub-Folder.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Object in Sub-Folder.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Report names contained in this Document. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFolder can be left as a blank string "", which will either return only an array of main level Report names (not located in folders), or if $bExhaustive is called with True, it will return an array of all Reports contained in the document.
;                  You can narrow the Report name return down to a specific folder by calling the appropriate path for the folder, separated by forward slashes (/), e.g. to get an array of Reports contained in folder 3, which is located in Folder 2, which is located inside folder 1, you would call $sFolder with the following path: Folder1/Folder2/Folder3
;                  All Reports located in folders will have the folder path prefixed to the Report name, separated by forward slashes (/). e.g. Folder1/Folder2/Folder3/ReportXYZ.
;                  Calling $bExhaustive with True when searching inside a Folder, will get all Report names from inside that folder, and all sub-folders.
;                  The order of the Report names inside the folders may not necessarily be in proper order, i.e. if there are two sub folders, and Folders inside the first sub-folder, the Reports inside the two folders will be listed first, then the Reports inside the folders inside the first sub-folder.
; Related .......: _LOBase_ReportsGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_ReportsGetNames(ByRef $oDoc, $bExhaustive = True, $sFolder = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj, $oSource
	Local Enum $iName, $iObj, $iPrefix
	Local $iCount = 0, $iFolders = 1, $iReports = 0
	Local $asNames[0], $asReports[0], $asFolderList[0], $asSplit[0]
	Local $avFolders[0][3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFolder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oSource = $oDoc.ReportDocuments()
	If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFolder <> "") Then
		$asSplit = StringSplit($sFolder, "/")

		For $i = 1 To $asSplit[0]
			If Not $oSource.hasByName($asSplit[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$oSource = $oSource.getByName($asSplit[$i])
			If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Next
		$sFolder &= "/"
	EndIf

	$asNames = $oSource.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	For $i = 0 To UBound($asNames) - 1
		$oObj = $oSource.getByName($asNames[$i])
		If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		If $oObj.supportsService("com.sun.star.ucb.Content") Then ; Form Doc.
			If (UBound($asReports) >= $iReports) Then ReDim $asReports[$iReports + 1]
			$asReports[$iReports] = $sFolder & $asNames[$i]
			$iReports += 1

		ElseIf $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
			ReDim $avFolders[1][3]
			$avFolders[0][$iName] = $asNames[$i]
			$avFolders[0][$iObj] = $oObj
			$avFolders[0][$iPrefix] = $asNames[$i] & "/"
		EndIf

		If $bExhaustive Then
			While ($iCount < UBound($avFolders))
				$asFolderList = $avFolders[$iCount][$iObj].getElementNames()
				If Not IsArray($asFolderList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				For $k = 0 To UBound($asFolderList) - 1
					$oObj = $avFolders[$iCount][$iObj].getByName($asFolderList[$k])
					If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

					If $oObj.supportsService("com.sun.star.ucb.Content") Then ; Form Doc.
						If (UBound($asReports) >= $iReports) Then ReDim $asReports[$iReports + 1]
						$asReports[$iReports] = $sFolder & $avFolders[$iCount][$iPrefix] & $asFolderList[$k]
						$iReports += 1

					ElseIf $oObj.supportsService("com.sun.star.sdb.Reports") Then ; Folder.
						ReDim $avFolders[$iFolders + 1][3]
						$avFolders[$iFolders][$iName] = $asFolderList[$k]
						$avFolders[$iFolders][$iObj] = $oObj
						$avFolders[$iFolders][$iPrefix] = $avFolders[$iCount][$iPrefix] & $asFolderList[$k] & "/"

						$iFolders += 1
					EndIf
				Next

				$iCount += 1
			WEnd

			If (UBound($avFolders) > 0) Then ReDim $avFolders[0][3]
			$iCount = 0
			$iFolders = 1
		EndIf
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($asReports), $asReports)
EndFunc   ;==>_LOBase_ReportsGetNames
