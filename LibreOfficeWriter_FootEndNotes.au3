#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter)
; AutoIt Version : v3.3.16.1
; UDF Version    : 0.0.0.3
; Description ...: Provides basic functionality through Autoit for interacting with Libre Office Writer.
; Author(s) .....: donnyh13, mLipok
; Sources .......: jguinch -- Printmgr.au3, used (_PrintMgr_EnumPrinter);
;					mLipok -- OOoCalc.au3, used (__OOoCalc_ComErrorHandler_UserFunction,_InternalComErrorHandler,
;						-- WriterDemo.au3, used _CreateStruct;
;					Andrew Pitonyak & Laurent Godard (VersionGet);
;					Leagnus & GMK -- OOoCalc.au3, used (SetPropertyValue)
; Dll ...........:
; Note...........: Tips/templates taken from OOoCalc UDF written by user GMK; also from Word UDF by user water.
;					I found the book by Andrew Pitonyak very helpful also, titled, "OpenOffice.org Macros Explained;
;						OOME Third Edition".
;					Of course, this UDF is written using the English version of LibreOffice, and may only work for the English
;						version of LibreOffice installations. Many functions in this UDF may or may not work with OpenOffice
;						Writer, however some settings are definitely for LibreOffice only.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_EndnoteDelete
; _LOWriter_EndnoteGetAnchor
; _LOWriter_EndnoteGetTextCursor
; _LOWriter_EndnoteInsert
; _LOWriter_EndnoteModifyAnchor
; _LOWriter_EndnoteSettingsAutoNumber
; _LOWriter_EndnoteSettingsStyles
; _LOWriter_EndnotesGetList
; _LOWriter_FootnoteDelete
; _LOWriter_FootnoteGetAnchor
; _LOWriter_FootnoteGetTextCursor
; _LOWriter_FootnoteInsert
; _LOWriter_FootnoteModifyAnchor
; _LOWriter_FootnoteSettingsAutoNumber
; _LOWriter_FootnoteSettingsContinuation
; _LOWriter_FootnoteSettingsStyles
; _LOWriter_FootnotesGetList
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteDelete
; Description ...: Delete a Endnote.
; Syntax ........: _LOWriter_EndnoteDelete(Byref $oEndNote)
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous _LOWriter_EndnoteInsert, or _LOWriter_EndnotesGetList function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oEndNote not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Endnote successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteDelete(ByRef $oEndNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oEndNote.dispose()
	$oEndNote = Null

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_EndnoteDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteGetAnchor
; Description ...: Create a Text Cursor at the Endnote Anchor position.
; Syntax ........: _LOWriter_EndnoteGetAnchor(Byref $oEndNote)
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous _LOWriter_EndnoteInsert, or _LOWriter_EndnotesGetList function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oEndNote not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully returned the Endnote Anchor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Anchor cursor returned is just a Text Cursor placed at the anchor's position.
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert, _LOWriter_CursorMove, _LOWriter_DocGetString, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteGetAnchor(ByRef $oEndNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnchor

	If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oAnchor = $oEndNote.Anchor.Text.createTextCursorByRange($oEndNote.Anchor())
	If Not IsObj($oAnchor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oAnchor)
EndFunc   ;==>_LOWriter_EndnoteGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteGetTextCursor
; Description ...: Create a Text Cursor in a Endnote to modify the text therein.
; Syntax ........: _LOWriter_EndnoteGetTextCursor(Byref $oEndNote)
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous _LOWriter_EndnoteInsert, or _LOWriter_EndnotesGetList function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oEndNote not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Cursor Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully retrieved the Endnote Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteGetTextCursor(ByRef $oEndNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oTextCursor = $oEndNote.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOWriter_EndnoteGetTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteInsert
; Description ...: Insert a Endnote into a Document.
; Syntax ........: _LOWriter_EndnoteInsert(Byref $oDoc, Byref $oCursor, $bOverwrite[, $sLabel = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten.
;				   +								If False, content will be inserted to the left of any selection.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the Endnote.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $oCursor is a Table cursor type, not supported.
;				   @Error 1 @Extended 5 Return 0 = $oCursor currently located in a Frame, Footnote, Endnote, or Header/ Footer cannot insert a Endnote in those data types.
;				   @Error 1 @Extended 6 Return 0 = $oCursor located in unknown data type.
;				   @Error 1 @Extended 7 Return 0 = $sLabel not a string.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 =  Error creating "com.sun.star.text.Endnote" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully inserted a new Endnote, returning Endnote Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Endnote cannot be inserted into a Frame, a Footnote, a Endnote, or the Header/ Footer.
; Related .......: _LOWriter_EndnoteDelete,  _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sLabel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oEndNote

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	Switch __LOWriter_Internal_CursorGetDataType($oDoc, $oCursor)

		Case $LOW_CURDATA_FRAME, $LOW_CURDATA_FOOTNOTE, $LOW_CURDATA_ENDNOTE, $LOW_CURDATA_HEADER_FOOTER
			Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0) ; Unsupported cursor type.
		Case $LOW_CURDATA_BODY_TEXT, $LOW_CURDATA_CELL
			$oEndNote = $oDoc.createInstance("com.sun.star.text.Endnote")
			If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

		Case Else
			Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0) ; Unknown Cursor type.
	EndSwitch

	If ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oEndNote.Label = $sLabel
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oEndNote, $bOverwrite)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oEndNote)
EndFunc   ;==>_LOWriter_EndnoteInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteModifyAnchor
; Description ...: Modify a Specific Endnote's settings.
; Syntax ........: _LOWriter_EndnoteModifyAnchor(Byref $oEndNote[, $sLabel = Null])
; Parameters ....: $oEndNote            - [in/out] an object. A Endnote Object from a previous _LOWriter_EndnoteInsert, or _LOWriter_EndnotesGetList function.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the Endnote. Set to "" for automatic numbering.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oEndNote not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sLabel not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = $sLabel was not set successfully.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Endnote settings were successfully modified.
;				   @Error 0 @Extended 1 Return String = Success. $sLabel set to Null, current Endnote Label returned.
;				   @Error 0 @Extended 2 Return String = Success. $sLabel set to Null, current Endnote AutoNumbering number returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteModifyAnchor(ByRef $oEndNote, $sLabel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oEndNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($sLabel = Null) Then
		; If Label is blank, return the AutoNumbering Number.
		If ($oEndNote.Label() = "") Then Return SetError($__LOW_STATUS_SUCCESS, 2, $oEndNote.Anchor.String())

		; Else return the Label.
		Return SetError($__LOW_STATUS_SUCCESS, 1, $oEndNote.Label())

	EndIf

	If Not IsString($sLabel) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oEndNote.Label = $sLabel
	If ($oEndNote.Label() <> $sLabel) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_EndnoteModifyAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteSettingsAutoNumber
; Description ...: Set or Retrieve Endnote Autonumbering settings.
; Syntax ........: _LOWriter_EndnoteSettingsAutoNumber(Byref $oDoc[, $iNumFormat = Null[, $iStartAt = Null[, $sBefore = Null[, $sAfter = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The numbering format to use for Endnote numbering. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iStartAt            - [optional] an integer value. Default is Null. The Number to begin Endnote counting from, Min. 1, Max 9999.
;                  $sBefore             - [optional] a string value. Default is Null. The text to display before a Endnote number in the note text.
;                  $sAfter              - [optional] a string value. Default is Null. The text to display after a Endnote number in the note text.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iNumFormat not an Integer, or Less than 0 or greater than 71. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iStartAt not an integer, less than 1 or greater than 9999.
;				   @Error 1 @Extended 4 Return 0 = $sBefore not a String.
;				   @Error 1 @Extended 5 Return 0 = $sAfter not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iNumFormat
;				   |								2 = Error setting $iStartAt
;				   |								4 = Error setting $sBefore
;				   |								8 = Error setting $sAfter
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_EndnotesGetList, _LOWriter_EndnoteInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteSettingsAutoNumber(ByRef $oDoc, $iNumFormat = Null, $iStartAt = Null, $sBefore = Null, $sAfter = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avENSettings[4]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iNumFormat, $iStartAt, $sBefore, $sAfter) Then
		__LOWriter_ArrayFill($avENSettings, $oDoc.EndnoteSettings.NumberingType(), ($oDoc.EndnoteSettings.StartAt() + 1), _
				$oDoc.EndnoteSettings.Prefix(), $oDoc.EndnoteSettings.Suffix())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avENSettings)
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDoc.EndnoteSettings.NumberingType = $iNumFormat
		$iError = ($oDoc.EndnoteSettings.NumberingType() = $iNumFormat) ? $iError : BitOR($iError, 1)
	EndIf

	; 0 Based -- Minus 1
	If ($iStartAt <> Null) Then
		If Not __LOWriter_IntIsBetween($iStartAt, 1, 9999) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDoc.EndnoteSettings.StartAt = ($iStartAt - 1)
		$iError = ($oDoc.EndnoteSettings.StartAt() = ($iStartAt - 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sBefore <> Null) Then
		If Not IsString($sBefore) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oDoc.EndnoteSettings.Prefix = $sBefore
		$iError = ($oDoc.EndnoteSettings.Prefix() = $sBefore) ? $iError : BitOR($iError, 4)
	EndIf

	If ($sAfter <> Null) Then
		If Not IsString($sAfter) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDoc.EndnoteSettings.Suffix = $sAfter
		$iError = ($oDoc.EndnoteSettings.Suffix() = $sAfter) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_EndnoteSettingsAutoNumber

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnoteSettingsStyles
; Description ...: Set or Retrieve Document Endnote Style settings.
; Syntax ........: _LOWriter_EndnoteSettingsStyles(Byref $oDoc[, $sParagraph = Null[, $sPage = Null[, $sTextArea = Null[, $sEndnoteArea = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sParagraph          - [optional] a string value. Default is Null. The Endnote Text Paragraph Style.
;                  $sPage               - [optional] a string value. Default is Null. The Page Style to use for the Endnote pages.
;                  $sTextArea           - [optional] a string value. Default is Null. The Character Style to use for the Endnote anchor in the document text.
;                  $sEndnoteArea        - [optional] a string value. Default is Null. The Character Style to use for the Endnote number in the Endnote text.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sParagraph not a String.
;				   @Error 1 @Extended 3 Return 0 = Paragraph Style referenced in $sParagraph not found in Document.
;				   @Error 1 @Extended 4 Return 0 = $sPage not a String.
;				   @Error 1 @Extended 5 Return 0 = Page Style referenced in $sPage not found in Document.
;				   @Error 1 @Extended 6 Return 0 = $sTextArea not a String.
;				   @Error 1 @Extended 7 Return 0 = Character Style referenced in $sTextArea not found in Document.
;				   @Error 1 @Extended 8 Return 0 = $sEndnoteArea not a String.
;				   @Error 1 @Extended 9 Return 0 = Character Style referenced in $sEndnoteArea not found in Document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sParagraph
;				   |								2 = Error setting $sPage
;				   |								4 = Error setting $sTextArea
;				   |								8 = Error setting $sEndnoteArea
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStylesGetNames, _LOWriter_CharStylesGetNames, _LOWriter_PageStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnoteSettingsStyles(ByRef $oDoc, $sParagraph = Null, $sPage = Null, $sTextArea = Null, $sEndnoteArea = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asENSettings[4]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sParagraph, $sPage, $sTextArea, $sEndnoteArea) Then
		__LOWriter_ArrayFill($asENSettings, __LOWriter_ParStyleNameToggle($oDoc.EndnoteSettings.ParaStyleName(), True), _
				__LOWriter_PageStyleNameToggle($oDoc.EndnoteSettings.PageStyleName(), True), _
				__LOWriter_CharStyleNameToggle($oDoc.EndnoteSettings.AnchorCharStyleName(), True), _
				__LOWriter_CharStyleNameToggle($oDoc.EndnoteSettings.CharStyleName(), True))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $asENSettings)
	EndIf

	If ($sParagraph <> Null) Then
		If Not IsString($sParagraph) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		If Not _LOWriter_ParStyleExists($oDoc, $sParagraph) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$sParagraph = __LOWriter_ParStyleNameToggle($sParagraph)
		$oDoc.EndnoteSettings.ParaStyleName = $sParagraph
		$iError = ($oDoc.EndnoteSettings.ParaStyleName() = $sParagraph) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sPage <> Null) Then
		If Not IsString($sPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		If Not _LOWriter_PageStyleExists($oDoc, $sPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$sPage = __LOWriter_PageStyleNameToggle($sPage)
		$oDoc.EndnoteSettings.PageStyleName = $sPage
		$iError = ($oDoc.EndnoteSettings.PageStyleName() = $sPage) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sTextArea <> Null) Then
		If Not IsString($sTextArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_CharStyleExists($oDoc, $sTextArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$sTextArea = __LOWriter_CharStyleNameToggle($sTextArea)
		$oDoc.EndnoteSettings.AnchorCharStyleName = $sTextArea
		$iError = ($oDoc.EndnoteSettings.AnchorCharStyleName() = $sTextArea) ? $iError : BitOR($iError, 4)
	EndIf

	If ($sEndnoteArea <> Null) Then
		If Not IsString($sEndnoteArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		If Not _LOWriter_CharStyleExists($oDoc, $sEndnoteArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$sEndnoteArea = __LOWriter_CharStyleNameToggle($sEndnoteArea)
		$oDoc.EndnoteSettings.CharStyleName = $sEndnoteArea
		$iError = ($oDoc.EndnoteSettings.CharStyleName() = $sEndnoteArea) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_EndnoteSettingsStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_EndnotesGetList
; Description ...: Retrieve an array of Endnote Objects contained in a Document.
; Syntax ........: _LOWriter_EndnotesGetList(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Endnotes Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully searched for Endnotes, none contained in document.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for Endnotes, Returning Array of Endnote Objects. @Extended set to number found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_EndnoteDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_EndnotesGetList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oEndNotes
	Local $aoEndnotes[0]
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oEndNotes = $oDoc.getEndnotes()
	If Not IsObj($oEndNotes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$iCount = $oEndNotes.getCount()

	If ($iCount > 0) Then
		ReDim $aoEndnotes[$iCount]

		For $i = 0 To $iCount - 1
			$aoEndnotes[$i] = $oEndNotes.getByIndex($i)

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return ($iCount > 0) ? SetError($__LOW_STATUS_SUCCESS, $iCount, $aoEndnotes) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_EndnotesGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteDelete
; Description ...: Delete a Footnote.
; Syntax ........: _LOWriter_FootnoteDelete(Byref $oFootNote)
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous _LOWriter_FootnoteInsert, Or _LOWriter_FootnotesGetList function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oFootNote not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Footnote successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FootnoteInsert, _LOWriter_FootnotesGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteDelete(ByRef $oFootNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oFootNote.dispose()
	$oFootNote = Null

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnoteDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteGetAnchor
; Description ...: Create a Text Cursor at the Footnote Anchor position.
; Syntax ........: _LOWriter_FootnoteGetAnchor(Byref $oFootNote)
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous _LOWriter_FootnoteInsert, Or _LOWriter_FootnotesGetList function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oFootNote not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully returned the Footnote Anchor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FootnotesGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteGetAnchor(ByRef $oFootNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAnchor

	If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oAnchor = $oFootNote.Anchor.Text.createTextCursorByRange($oFootNote.Anchor())
	If Not IsObj($oAnchor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oAnchor)
EndFunc   ;==>_LOWriter_FootnoteGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteGetTextCursor
; Description ...: Create a Text Cursor in a Footnote to modify the text therein.
; Syntax ........: _LOWriter_FootnoteGetTextCursor(Byref $oFootNote)
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous _LOWriter_FootnoteInsert, Or _LOWriter_FootnotesGetList function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oFootNote not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Cursor Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully retrieved the footnote Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_CursorMove, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteGetTextCursor(ByRef $oFootNote)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oTextCursor = $oFootNote.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOWriter_FootnoteGetTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteInsert
; Description ...: Insert a Footnote into a Document.
; Syntax ........: _LOWriter_FootnoteInsert(Byref $oDoc, Byref $oCursor, $bOverwrite[, $sLabel = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten.
;				   +									If False, content will be inserted to the left of any selection.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the footnote.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $oCursor is a Table cursor type, not supported.
;				   @Error 1 @Extended 5 Return 0 = $oCursor currently located in a Frame, Footnote, Endnote, or Header/Footer,
;				   +									cannot insert a Footnote in those data types.
;				   @Error 1 @Extended 6 Return 0 = $oCursor located in unknown data type.
;				   @Error 1 @Extended 7 Return 0 = $sLabel not a string.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 =  Error creating "com.sun.star.text.Footnote" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully inserted a new footnote, returning Footnote
;				   +									Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Footnote cannot be inserted into a Frame, a Footnote, a Endnote, or a Header/ Footer.
; Related .......: _LOWriter_FootnoteDelete, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sLabel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFootNote

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	Switch __LOWriter_Internal_CursorGetDataType($oDoc, $oCursor)

		Case $LOW_CURDATA_FRAME, $LOW_CURDATA_FOOTNOTE, $LOW_CURDATA_ENDNOTE, $LOW_CURDATA_HEADER_FOOTER
			Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0) ; Unsupported cursor type.
		Case $LOW_CURDATA_BODY_TEXT, $LOW_CURDATA_CELL
			$oFootNote = $oDoc.createInstance("com.sun.star.text.Footnote")
			If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

		Case Else
			Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0) ; Unknown Cursor type.
	EndSwitch

	If ($sLabel <> Null) Then
		If Not IsString($sLabel) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oFootNote.Label = $sLabel
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oFootNote, $bOverwrite)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oFootNote)
EndFunc   ;==>_LOWriter_FootnoteInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteModifyAnchor
; Description ...: Modify a Footnote's Anchor Character.
; Syntax ........: _LOWriter_FootnoteModifyAnchor(Byref $oFootNote[, $sLabel = Null])
; Parameters ....: $oFootNote           - [in/out] an object. A Footnote Object from a previous _LOWriter_FootnoteInsert, Or _LOWriter_FootnotesGetList function.
;                  $sLabel              - [optional] a string value. Default is Null. A custom anchor label for the Footnote. Set to "" for automatic numbering.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oFootNote not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sLabel not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Failed to set $sLabel.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Footnote settings were successfully modified.
;				   @Error 0 @Extended 1 Return String = Success. $sLabel set to Null, current Footnote Custom Label returned.
;				   @Error 0 @Extended 2 Return String = Success. $sLabel set to Null, current Footnote AutoNumbering number returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FootnoteInsert, _LOWriter_FootnotesGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteModifyAnchor(ByRef $oFootNote, $sLabel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oFootNote) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($sLabel = Null) Then
		; If Label is blank, return the AutoNumbering Number.
		If ($oFootNote.Label() = "") Then Return SetError($__LOW_STATUS_SUCCESS, 2, $oFootNote.Anchor.String())

		; Else return the Label.
		Return SetError($__LOW_STATUS_SUCCESS, 1, $oFootNote.Label())
	EndIf

	If Not IsString($sLabel) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oFootNote.Label = $sLabel
	If ($oFootNote.Label() <> $sLabel) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnoteModifyAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteSettingsAutoNumber
; Description ...: Set or Retrieve Footnote Autonumbering settings.
; Syntax ........: _LOWriter_FootnoteSettingsAutoNumber(Byref $oDoc[, $iNumFormat = Null[, $iStartAt = Null[, $sBefore = Null[, $sAfter = Null[, $iCounting = Null[, $bEndOfDoc = Null]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The numbering format to use for Footnote numbering. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iStartAt            - [optional] an integer value. Default is Null. The Number to begin Footnote counting from, this is labeled "Counting" in the L.O. User Interface. Min. 1, Max 9999.
;                  $sBefore             - [optional] a string value. Default is Null. The text to display before a Footnote number in the note text.
;                  $sAfter              - [optional] a string value. Default is Null. The text to display after a Footnote number in the note text.
;                  $iCounting           - [optional] an integer value (0-2). Default is Null. The Counting type of the footnotes, such as per page etc., see constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3..
;                  $bEndOfDoc           - [optional] a boolean value. Default is Null. If True, Footnotes are placed at the end of the document, like Endnotes.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iNumFormat not an Integer, Less than 0 or greater than 71. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iStartAt not an integer, less than 1 or greater than 9999.
;				   @Error 1 @Extended 4 Return 0 = $sBefore not a String.
;				   @Error 1 @Extended 5 Return 0 = $sAfter not a String.
;				   @Error 1 @Extended 6 Return 0 = $iCounting not an Integer, less than 0 or greater than 2. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3..
;				   @Error 1 @Extended 7 Return 0 = $bEndOfDoc not a boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iNumFormat
;				   |								2 = Error setting $iStartAt
;				   |								4 = Error setting $sBefore
;				   |								8 = Error setting $sAfter
;				   |								16 = Error setting $iCounting
;				   |								32 = Error setting $bEndOfDoc
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteSettingsAutoNumber(ByRef $oDoc, $iNumFormat = Null, $iStartAt = Null, $sBefore = Null, $sAfter = Null, $iCounting = Null, $bEndOfDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFNSettings[6]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iNumFormat, $iStartAt, $sBefore, $sAfter, $iCounting, $bEndOfDoc) Then
		__LOWriter_ArrayFill($avFNSettings, $oDoc.FootnoteSettings.NumberingType(), ($oDoc.FootnoteSettings.StartAt + 1), _
				$oDoc.FootnoteSettings.Prefix(), $oDoc.FootnoteSettings.Suffix(), $oDoc.FootnoteSettings.FootnoteCounting(), _
				$oDoc.FootnoteSettings.PositionEndOfDoc())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avFNSettings)
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDoc.FootnoteSettings.NumberingType = $iNumFormat
		$iError = ($oDoc.FootnoteSettings.NumberingType() = $iNumFormat) ? $iError : BitOR($iError, 1)
	EndIf

	; 0 Based -- Minus 1
	If ($iStartAt <> Null) Then
		If Not __LOWriter_IntIsBetween($iStartAt, 1, 9999) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDoc.FootnoteSettings.StartAt = ($iStartAt - 1)
		$iError = ($oDoc.FootnoteSettings.StartAt() = ($iStartAt - 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sBefore <> Null) Then
		If Not IsString($sBefore) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oDoc.FootnoteSettings.Prefix = $sBefore
		$iError = ($oDoc.FootnoteSettings.Prefix() = $sBefore) ? $iError : BitOR($iError, 4)
	EndIf

	If ($sAfter <> Null) Then
		If Not IsString($sAfter) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDoc.FootnoteSettings.Suffix = $sAfter
		$iError = ($oDoc.FootnoteSettings.Suffix() = $sAfter) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iCounting <> Null) Then
		If Not __LOWriter_IntIsBetween($iCounting, $LOW_FOOTNOTE_COUNT_PER_PAGE, $LOW_FOOTNOTE_COUNT_PER_DOC) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oDoc.FootnoteSettings.FootnoteCounting = $iCounting
		$iError = ($oDoc.FootnoteSettings.FootnoteCounting() = $iCounting) ? $iError : BitOR($iError, 16)
	EndIf

	If ($bEndOfDoc <> Null) Then
		If Not IsBool($bEndOfDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oDoc.FootnoteSettings.PositionEndOfDoc = $bEndOfDoc
		$iError = ($oDoc.FootnoteSettings.PositionEndOfDoc() = $bEndOfDoc) ? $iError : BitOR($iError, 32)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnoteSettingsAutoNumber

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteSettingsContinuation
; Description ...: Set or Retrieve Footnote continuation settings.
; Syntax ........: _LOWriter_FootnoteSettingsContinuation(Byref $oDoc[, $sEnd = Null[, $sBegin = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sEnd                - [optional] a string value. Default is Null. The text to display at the end of a Footnote before it continues on the next page.
;                  $sBegin              - [optional] a string value. Default is Null. The text to display at the beginning of a Footnote that has continued on the next page.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sEnd not a String.
;				   @Error 1 @Extended 3 Return 0 = $sBegin not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sEnd
;				   |								2 = Error setting $sBegin
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteSettingsContinuation(ByRef $oDoc, $sEnd = Null, $sBegin = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asFNSettings[2]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sEnd, $sBegin) Then
		__LOWriter_ArrayFill($asFNSettings, $oDoc.FootnoteSettings.EndNotice(), $oDoc.FootnoteSettings.BeginNotice())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $asFNSettings)
	EndIf

	If ($sEnd <> Null) Then
		If Not IsString($sEnd) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDoc.FootnoteSettings.EndNotice = $sEnd
		$iError = ($oDoc.FootnoteSettings.EndNotice() = $sEnd) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sBegin <> Null) Then
		If Not IsString($sBegin) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDoc.FootnoteSettings.BeginNotice = $sBegin
		$iError = ($oDoc.FootnoteSettings.BeginNotice() = $sBegin) ? $iError : BitOR($iError, 2)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnoteSettingsContinuation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnoteSettingsStyles
; Description ...: Set or Retrieve Document Footnote Style settings.
; Syntax ........: _LOWriter_FootnoteSettingsStyles(Byref $oDoc[, $sParagraph = Null[, $sPage = Null[, $sTextArea = Null[, $sFootnoteArea = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sParagraph          - [optional] a string value. Default is Null. The Footnote Text Paragraph Style.
;                  $sPage               - [optional] a string value. Default is Null. The Page Style to use for the Footnote pages. Only valid if the footnotes are set to End of Document, instead of per page.
;                  $sTextArea           - [optional] a string value. Default is Null. The Character Style to use for the Footnote anchor in the document text.
;                  $sFootnoteArea       - [optional] a string value. Default is Null. The Character Style to use for the Footnote number in the footnote text.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sParagraph not a String.
;				   @Error 1 @Extended 3 Return 0 = Paragraph Style referenced in $sParagraph not found in Document.
;				   @Error 1 @Extended 4 Return 0 = $sPage not a String.
;				   @Error 1 @Extended 5 Return 0 = Page Style referenced in $sPage not found in Document.
;				   @Error 1 @Extended 6 Return 0 = $sTextArea not a String.
;				   @Error 1 @Extended 7 Return 0 = Character Style referenced in $sTextArea not found in Document.
;				   @Error 1 @Extended 8 Return 0 = $sFootnoteArea not a String.
;				   @Error 1 @Extended 9 Return 0 = Character Style referenced in $sFootnoteArea not found in Document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sParagraph
;				   |								2 = Error setting $sPage
;				   |								4 = Error setting $sTextArea
;				   |								8 = Error setting $sFootnoteArea
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStylesGetNames, _LOWriter_PageStylesGetNames, _LOWriter_CharStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnoteSettingsStyles(ByRef $oDoc, $sParagraph = Null, $sPage = Null, $sTextArea = Null, $sFootnoteArea = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFNSettings[4]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sParagraph, $sPage, $sTextArea, $sFootnoteArea) Then
		__LOWriter_ArrayFill($avFNSettings, __LOWriter_ParStyleNameToggle($oDoc.FootnoteSettings.ParaStyleName(), True), _
				__LOWriter_PageStyleNameToggle($oDoc.FootnoteSettings.PageStyleName(), True), _
				__LOWriter_CharStyleNameToggle($oDoc.FootnoteSettings.AnchorCharStyleName(), True), _
				__LOWriter_CharStyleNameToggle($oDoc.FootnoteSettings.CharStyleName(), True))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avFNSettings)
	EndIf

	If ($sParagraph <> Null) Then
		If Not IsString($sParagraph) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		If Not _LOWriter_ParStyleExists($oDoc, $sParagraph) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$sParagraph = __LOWriter_ParStyleNameToggle($sParagraph)
		$oDoc.FootnoteSettings.ParaStyleName = $sParagraph
		$iError = ($oDoc.FootnoteSettings.ParaStyleName() = $sParagraph) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sPage <> Null) Then
		If Not IsString($sPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		If Not _LOWriter_PageStyleExists($oDoc, $sPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$sPage = __LOWriter_PageStyleNameToggle($sPage)
		$oDoc.FootnoteSettings.PageStyleName = $sPage
		$iError = ($oDoc.FootnoteSettings.PageStyleName() = $sPage) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sTextArea <> Null) Then
		If Not IsString($sTextArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_CharStyleExists($oDoc, $sTextArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$sTextArea = __LOWriter_CharStyleNameToggle($sTextArea)
		$oDoc.FootnoteSettings.AnchorCharStyleName = $sTextArea
		$iError = ($oDoc.FootnoteSettings.AnchorCharStyleName() = $sTextArea) ? $iError : BitOR($iError, 4)
	EndIf

	If ($sFootnoteArea <> Null) Then
		If Not IsString($sFootnoteArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		If Not _LOWriter_CharStyleExists($oDoc, $sFootnoteArea) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$sFootnoteArea = __LOWriter_CharStyleNameToggle($sFootnoteArea)
		$oDoc.FootnoteSettings.CharStyleName = $sFootnoteArea
		$iError = ($oDoc.FootnoteSettings.CharStyleName() = $sFootnoteArea) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnoteSettingsStyles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FootnotesGetList
; Description ...: Retrieve an array of Footnote Objects contained in a Document.
; Syntax ........: _LOWriter_FootnotesGetList(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Footnotes Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully searched for Footnotes, none contained in document.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for Footnotes, Returning Array of Footnote Objects. @Extended set to number found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FootnoteDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FootnotesGetList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFootNotes
	Local $aoFootnotes[0]
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oFootNotes = $oDoc.getFootnotes()
	If Not IsObj($oFootNotes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$iCount = $oFootNotes.getCount()

	If ($iCount > 0) Then
		ReDim $aoFootnotes[$iCount]

		For $i = 0 To $iCount - 1
			$aoFootnotes[$i] = $oFootNotes.getByIndex($i)
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return ($iCount > 0) ? SetError($__LOW_STATUS_SUCCESS, $iCount, $aoFootnotes) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FootnotesGetList
