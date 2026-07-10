#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
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
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Closing, Saving, etc. L.O. Calc documents.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_DocClose
; _LOCalc_DocColumnsRowsAreFrozen
; _LOCalc_DocColumnsRowsFreeze
; _LOCalc_DocConnect
; _LOCalc_DocCreate
; _LOCalc_DocExport
; _LOCalc_DocFormulaBarHeight
; _LOCalc_DocGetName
; _LOCalc_DocGetPath
; _LOCalc_DocHasPath
; _LOCalc_DocIsActive
; _LOCalc_DocIsModified
; _LOCalc_DocIsReadOnly
; _LOCalc_DocMaximize
; _LOCalc_DocMinimize
; _LOCalc_DocOpen
; _LOCalc_DocPosAndSize
; _LOCalc_DocPrint
; _LOCalc_DocRedo
; _LOCalc_DocRedoClear
; _LOCalc_DocRedoCurActionTitle
; _LOCalc_DocRedoGetAllActionTitles
; _LOCalc_DocRedoIsPossible
; _LOCalc_DocSave
; _LOCalc_DocSaveAs
; _LOCalc_DocSelectionCopy
; _LOCalc_DocSelectionGet
; _LOCalc_DocSelectionPaste
; _LOCalc_DocSelectionSet
; _LOCalc_DocSelectionSetMulti
; _LOCalc_DocToFront
; _LOCalc_DocUndo
; _LOCalc_DocUndoActionBegin
; _LOCalc_DocUndoActionEnd
; _LOCalc_DocUndoClear
; _LOCalc_DocUndoCurActionTitle
; _LOCalc_DocUndoGetAllActionTitles
; _LOCalc_DocUndoIsPossible
; _LOCalc_DocUndoReset
; _LOCalc_DocViewDisplaySettings
; _LOCalc_DocViewWindowSettings
; _LOCalc_DocVisible
; _LOCalc_DocWindowFirstColumn
; _LOCalc_DocWindowFirstRow
; _LOCalc_DocWindowIsSplit
; _LOCalc_DocWindowSplit
; _LOCalc_DocWindowVisibleRange
; _LOCalc_DocZoom
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocClose
; Description ...: Close an existing Calc Document, returning its save path if applicable.
; Syntax ........: _LOCalc_DocClose(ByRef $oDoc[, $bSaveChanges = True[, $sSaveName = ""[, $bDeliverOwnership = True]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bSaveChanges        - [optional] Default is True. If True, saves changes if any were made before closing. See remarks.
;                  $sSaveName           - [optional] Default is "". The file name to save the file as, if the file hasn't been saved before. See Remarks.
;                  $bDeliverOwnership   - [optional] Default is True. If True, deliver ownership of the document Object from the script to LibreOffice, recommended is True.
; Return values .: Success: String
;                  @Error: 0, @Extended: 1, Return: String = Success, Document was successfully closed, and was saved to the returned file Path.
;                  @Error: 0, @Extended: 2, Return: String = Success, Document was successfully closed, document's changes were saved to its existing location.
;                  @Error: 0, @Extended: 3, Return: String = Success, Document was successfully closed, document either had no changes to save, or $bSaveChanges was called with False. If document had a save location, or if document was saved to a location, it is returned, else an empty string is returned.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bSaveChanges not a Boolean.
;                  @Error: 1, @Extended: 3 = $sSaveName not a String.
;                  @Error: 1, @Extended: 4 = $bDeliverOwnership not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error while creating Filter Name properties.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Path Conversion to L.O. URL Failed.
;                  @Error: 3, @Extended: 2 = Error while retrieving FilterName.
;                  @Error: 3, @Extended: 3 = Failed to close Document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bSaveChanges is True and the document hasn't been saved yet, the document is saved to the desktop.
;                  If $sSaveName is undefined, it is saved as an .ods document to the desktop, named Year-Month-Day_Hour-Minute-Second.ods. $sSaveName may be a name only without an extension, in which case the file will be saved in .ods format. Or you may define your own format by including an extension, such as "Test.xlsx"
; Related .......: _LOCalc_DocOpen, _LOCalc_DocConnect, _LOCalc_DocCreate, _LOCalc_DocSaveAs, _LOCalc_DocSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocClose(ByRef $oDoc, $bSaveChanges = True, $sSaveName = "", $bDeliverOwnership = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sDocPath = "", $sSavePath, $sFilterName
	Local $aArgs[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bSaveChanges) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSaveName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDeliverOwnership) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If Not $oDoc.hasLocation() And ($bSaveChanges = True) Then
		$sSavePath = @DesktopDir & "\"
		If ($sSaveName = "") Or ($sSaveName = " ") Then
			$sSaveName = @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "-" & @MIN & "-" & @SEC & ".ods"
			$sFilterName = "calc8"
		EndIf

		$sSavePath = _LO_PathConvert($sSavePath & $sSaveName, 1)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If $sFilterName = "" Then $sFilterName = __LOCalc_FilterNameGet($sSavePath)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$aArgs[0] = __LO_SetPropertyValue("FilterName", $sFilterName)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	If ($bSaveChanges = True) Then
		If $oDoc.hasLocation() Then
			$oDoc.store()
			$sDocPath = _LO_PathConvert($oDoc.getURL(), $LO_PATHCONV_PCPATH_RETURN)
			$oDoc.Close($bDeliverOwnership)

			If Not __LO_IsObjInvalid($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$oDoc = Null

			Return SetError($__LO_STATUS_SUCCESS, 2, $sDocPath)

		Else
			$oDoc.storeAsURL($sSavePath, $aArgs)
			$oDoc.Close($bDeliverOwnership)

			If Not __LO_IsObjInvalid($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$oDoc = Null

			Return SetError($__LO_STATUS_SUCCESS, 1, _LO_PathConvert($sSavePath, $LO_PATHCONV_PCPATH_RETURN))
		EndIf
	EndIf

	If $oDoc.hasLocation() Then $sDocPath = _LO_PathConvert($oDoc.getURL(), $LO_PATHCONV_PCPATH_RETURN)

	$oDoc.Close($bDeliverOwnership)

	If Not __LO_IsObjInvalid($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oDoc = Null

	Return SetError($__LO_STATUS_SUCCESS, 3, $sDocPath)
EndFunc   ;==>_LOCalc_DocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocColumnsRowsAreFrozen
; Description ...: Query whether the Document has Columns or Rows currently frozen in view.
; Syntax ........: _LOCalc_DocColumnsRowsAreFrozen(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Document currently contains frozen Columns/Rows.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query Document whether frozen Columns/Rows are present.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocColumnsRowsFreeze
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocColumnsRowsAreFrozen(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bReturn = $oDoc.CurrentController.hasFrozenPanes()
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOCalc_DocColumnsRowsAreFrozen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocColumnsRowsFreeze
; Description ...: Set Columns and/or Rows of a document to be frozen in view.
; Syntax ........: _LOCalc_DocColumnsRowsFreeze(ByRef $oDoc[, $iColumns = 0[, $iRows = 0]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iColumns            - [optional] Default is 0. The number of Columns to freeze. Call with 0 to skip. See remarks.
;                  $iRows               - [optional] Default is 0. The number of Rows to freeze. See remarks.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Called Columns/Rows were successfully frozen.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iColumns not an Integer, less than 0 or greater than number of columns contained in the document.
;                  @Error: 1, @Extended: 3 = $iRows not an Integer, less than 0 or greater than number of rows contained in the document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To set no Columns or rows to be frozen in view, set both $iColumns and $iRows to 0.
;                  Setting either $iColumns or $iRows will lose previous values for both.
; Related .......: _LOCalc_DocColumnsRowsAreFrozen
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocColumnsRowsFreeze(ByRef $oDoc, $iColumns = 0, $iRows = 0)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iColumns, 0, $oDoc.CurrentController.getActiveSheet().Columns.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iRows, 0, $oDoc.CurrentController.getActiveSheet().Rows.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oDoc.CurrentController.freezeAtPosition($iColumns, $iRows)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocColumnsRowsFreeze

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocConnect
; Description ...: Connect to an already opened instance of LibreOffice Calc.
; Syntax ........: _LOCalc_DocConnect([$iMode = $LO_DOC_CONNECT_MODE_CURRENT[, $sSearch = ""[, $bCaseless = False]]])
; Parameters ....: $iMode               - [optional] (0-4) Default is $LO_DOC_CONNECT_MODE_CURRENT. The Connect mode. See Constants, $LO_DOC_CONNECT_MODE_* as defined in LibreOffice_Constants.au3.
;                  $sSearch             - [optional] Default is "". The Name, Title or Path of the Document to search for. See remarks.
;                  $bCaseless           - [optional] Default is False. If True, searches are caseless when using $LO_DOC_CONNECT_MODE_SEARCH_* flags.
; Return values .: Success: Object or Array.
;                  @Error: 0, @Extended: 1, Return: Object = Success, The Object for the current, or last active Calc document is returned.
;                  @Error: 0, @Extended: 1, Return: Object = Success, The Object for the found Document with matching Name, Title or Path.
;                  @Error: 0, @Extended: ?, Return: Array = Success, An Array of all open LibreOffice Calc Documents. @Extended is set to number of results. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $iMode not an Integer, less than 0 or greater than 4. See Constants, $LO_DOC_CONNECT_MODE_* as defined in LibreOffice_Constants.au3.
;                  @Error: 1, @Extended: 2 = $sSearch not a String.
;                  @Error: 1, @Extended: 3 = $bCaseless not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating ServiceManager object.
;                  @Error: 2, @Extended: 2 = Error creating Desktop object.
;                  @Error: 2, @Extended: 3 = Error creating enumeration of open documents.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = No open LibreOffice documents.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Document Object.
;                  @Error: 3, @Extended: 3 = Failed to identify Document type.
;                  @Error: 3, @Extended: 4 = Error converting path to LibreOffice URL.
;                  @Error: 3, @Extended: 5 = Current Document not a Calc Document.
;                  @Error: 3, @Extended: 6 = No matches found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only Calc documents are searched or returned using any of the flags.
;                  The value used for $sSearch depends on the flag called in $iMode. It is ignored except for the $LO_DOC_CONNECT_MODE_SEARCH_* flags.
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_TITLE, $sSearch must be the full Title with Office and Component name; e.g: "Test.ods — LibreOffice Calc". This will be the same Title AutoIt would match or return from functions like WinGetTitle.
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_NAME, $sSearch must be the Document's full name, without the extension; e.g: "Test".
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_NAME_WITH_EXT, $sSearch must be the Document's name, with the extension; e.g: "Test.ods". If the Document hasn't been saved, just the name will work, e.g., "Untitled 1".
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_PATH, $sSearch must be the full Path of the document (Name and extension included); e.g: "C:\file\Test.ods."
;                  The Connect All option returns a single columned array. ($aArray[0]), each result is stored in a separate row.
;                  -Row 1 contains the Object for that document. e.g. $aArray[0] = $oDoc
;                  -Row 2 contains the Object for the next document. e.g. $aArray[1] = $oDoc2. And so on.
; Related .......: _LOCalc_DocOpen, _LOCalc_DocClose, _LOCalc_DocCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocConnect($iMode = $LO_DOC_CONNECT_MODE_CURRENT, $sSearch = "", $bCaseless = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0, $iDocType
	Local $aoConnectAll[0]
	Local $sCaseless = ""
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop

	If Not __LO_IntIsBetween($iMode, $LO_DOC_CONNECT_MODE_ALL, $LO_DOC_CONNECT_MODE_SEARCH_PATH) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sSearch) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bCaseless) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	If Not $oDesktop.getComponents.hasElements() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; no L.O open

	Switch $iMode
		Case $LO_DOC_CONNECT_MODE_ALL
			$oEnumDoc = $oDesktop.getComponents.createEnumeration()
			If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

			While $oEnumDoc.hasMoreElements()
				$oDoc = $oEnumDoc.nextElement()
				If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				$iDocType = _LO_DocGetType($oDoc)
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.

				If ($iDocType = $LO_DOC_TYPE_CALC) Then
					If (UBound($aoConnectAll) <= $iCount) Then ReDim $aoConnectAll[$iCount + 1]
					$aoConnectAll[$iCount] = $oDoc
					$iCount += 1
				EndIf
				Sleep((IsInt($iCount / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
			WEnd

			Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoConnectAll)

		Case $LO_DOC_CONNECT_MODE_CURRENT
			$oDoc = $oDesktop.currentComponent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$iDocType = _LO_DocGetType($oDoc)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.
			If ($iDocType <> $LO_DOC_TYPE_CALC) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; Not a Calc Doc.

			Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc)

		Case $LO_DOC_CONNECT_MODE_SEARCH_TITLE, $LO_DOC_CONNECT_MODE_SEARCH_NAME, $LO_DOC_CONNECT_MODE_SEARCH_NAME_WITH_EXT, $LO_DOC_CONNECT_MODE_SEARCH_PATH
			$sSearch = StringRegExpReplace($sSearch, "(^\s*|\s*$)", "") ; Strip leading and trailing spaces

			If $bCaseless Then $sCaseless = "(?i)"

			If ($iMode = $LO_DOC_CONNECT_MODE_SEARCH_PATH) Then
				$sSearch = _LO_PathConvert($sSearch, $LO_PATHCONV_OFFICE_RETURN) ; Convert to L.O File path.
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
			EndIf

			$oEnumDoc = $oDesktop.getComponents.createEnumeration()
			If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

			While $oEnumDoc.hasMoreElements()
				$oDoc = $oEnumDoc.nextElement()
				If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				$iDocType = _LO_DocGetType($oDoc)
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.

				If ($iDocType = $LO_DOC_TYPE_CALC) Then
					Switch $iMode
						Case $LO_DOC_CONNECT_MODE_SEARCH_TITLE
							; First make sure Current Controller is available (It wont be if Document is opened Hidden, in some Components.).
							If IsObj($oDoc.CurrentController()) And StringRegExp($oDoc.CurrentController.Frame.Title(), $sCaseless & "\Q" & $sSearch & "\E") Then

								Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc)
							EndIf

						Case $LO_DOC_CONNECT_MODE_SEARCH_NAME
							; Allow space(s) after name in case user put some in the Document name.
							; Add additional capture for Extension to just match the name the user put in, else force match at end of String for unsaved Documents.
							If StringRegExp($oDoc.Title(), $sCaseless & "\Q" & $sSearch & "\E\s*(\.\w+)?$") Then

								Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc)
							EndIf

						Case $LO_DOC_CONNECT_MODE_SEARCH_NAME_WITH_EXT
							If StringRegExp($oDoc.Title(), $sCaseless & "\Q" & $sSearch & "\E") Then

								Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc)
							EndIf

						Case $LO_DOC_CONNECT_MODE_SEARCH_PATH
							If StringRegExp($oDoc.getURL(), $sCaseless & "\Q" & $sSearch & "\E") Then

								Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc)
							EndIf
					EndSwitch
				EndIf
			WEnd
	EndSwitch

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0) ; No matches
EndFunc   ;==>_LOCalc_DocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocCreate
; Description ...: Open a new LibreOffice Calc Document or Connect to an existing blank, unsaved, writable document.
; Syntax ........: _LOCalc_DocCreate([$bForceNew = True[, $bHidden = False]])
; Parameters ....: $bForceNew           - [optional] Default is True. If True, force opening a new Calc Document instead of checking for a usable blank.
;                  $bHidden             - [optional] Default is False. If True opens the new document invisible or changes the existing document to invisible.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 1, Return: Object = Successfully connected to an existing Document. Returning Document's Object
;                  @Error: 0, @Extended: 2, Return: Object = Successfully created a new document. Returning Document's Object
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $bForceNew not a Boolean.
;                  @Error: 1, @Extended: 2 = $bHidden not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failure Creating Object com.sun.star.ServiceManager.
;                  @Error: 2, @Extended: 2 = Failure Creating Object com.sun.star.frame.Desktop.
;                  @Error: 2, @Extended: 3 = Failed to enumerate available documents.
;                  @Error: 2, @Extended: 4 = Failure Creating New Document.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Document Object is still returned. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHidden
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocOpen, _LOCalc_DocClose, _LOCalc_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocCreate($bForceNew = True, $bHidden = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $aArgs[1]
	Local $iError = 0
	Local $oServiceManager, $oDesktop, $oDoc, $oEnumDoc
	Local $sServiceName = "com.sun.star.sheet.SpreadsheetDocument"

	If Not IsBool($bForceNew) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$aArgs[0] = __LO_SetPropertyValue("Hidden", $bHidden)
	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; If not force new, and L.O pages exist then see if there are any blank calc documents to use.
	If Not $bForceNew And $oDesktop.getComponents.hasElements() Then
		$oEnumDoc = $oDesktop.getComponents.createEnumeration()
		If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService($sServiceName) _
					And Not ($oDoc.hasLocation() And Not $oDoc.isReadOnly()) And Not ($oDoc.isModified() = 0) Then
				$oDoc.CurrentController.Frame.ContainerWindow.Visible = ($bHidden) ? (False) : (True) ; opposite value of $bHidden.
				$iError = ($oDoc.CurrentController.Frame.isHidden() = $bHidden) ? ($iError) : (BitOR($iError, 1))

				Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))
			EndIf
		WEnd
	EndIf

	If Not IsObj($aArgs[0]) Then $iError = BitOR($iError, 1)
	$oDoc = $oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", $iURLFrameCreate, $aArgs)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOCalc_DocCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocExport
; Description ...: Export a Document with the specified file name to the path specified, with any parameters used.
; Syntax ........: _LOCalc_DocExport(ByRef $oDoc, $sFilePath[, $bSamePath = False[, $sFilterName = ""[, $bOverwrite = Null[, $sPassword = Null]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sFilePath           - Full path to save the document to, including Filename and extension. See Remarks.
;                  $bSamePath           - [optional] Default is False. If True, uses the path of the current document to export to. See Remarks
;                  $sFilterName         - [optional] Default is "". Filter name. If called with "" (blank string), Filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .ods extension is used instead, with the filter name of "calc8".
;                  $bOverwrite          - [optional] Default is Null. If True, file will be overwritten.
;                  $sPassword           - [optional] Default is Null. Password String to set for the document. (Not all file formats can have a Password set). "" (blank string) or Null = No Password.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Success. Returning save path for exported document.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sFilePath not a String.
;                  @Error: 1, @Extended: 3 = $bSamePath not a Boolean.
;                  @Error: 1, @Extended: 4 = $sFilterName not a String.
;                  @Error: 1, @Extended: 5 = $bOverwrite not a Boolean.
;                  @Error: 1, @Extended: 6 = $sPassword not a String.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating FilterName Property
;                  @Error: 2, @Extended: 2 = Error creating Overwrite Property
;                  @Error: 2, @Extended: 3 = Error creating Password Property
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error Converting Path to/from L.O. URL
;                  @Error: 3, @Extended: 2 = Document has no save path, and $bSamePath is called with True.
;                  @Error: 3, @Extended: 3 = Error retrieving FilterName.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Does not alter the original save path (if there was one), saves a copy of the document to the new path, in the new file format if one is chosen.
;                  If $bSamePath is called with True, the same save path as the current document is used. You must still fill in "$sFilePath" with the desired File Name and new extension, but you do not need to enter the file path.
; Related .......: _LOCalc_DocSave, _LOCalc_DocSaveAs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocExport(ByRef $oDoc, $sFilePath, $bSamePath = False, $sFilterName = "", $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aProperties[3]
	Local $sOriginalPath, $sSavePath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bSamePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sFilterName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If $bSamePath Then
		If $oDoc.hasLocation() Then
			$sOriginalPath = $oDoc.getURL()
			$sOriginalPath = StringLeft($sOriginalPath, StringInStr($sOriginalPath, "/", 0, -1)) ; Cut the original name off.
			If StringInStr($sFilePath, "\") Then $sFilePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_OFFICE_RETURN) ; Convert to L.O. URL
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$sFilePath = $sOriginalPath & $sFilePath ; combine the path with the new name.

		Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndIf
	EndIf

	If Not $bSamePath Then $sFilePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFilterName = "") Or ($sFilterName = " ") Then $sFilterName = __LOCalc_FilterNameGet($sFilePath, True)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$aProperties[0] = __LO_SetPropertyValue("FilterName", $sFilterName)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	EndIf

	If ($sPassword <> Null) Then
		If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)
	EndIf

	$oDoc.storeToURL($sFilePath, $aProperties)

	$sSavePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOCalc_DocExport

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocFormulaBarHeight
; Description ...: Set or Retrieve the current Formula Bar Height. L.O. 7.4+
; Syntax ........: _LOCalc_DocFormulaBarHeight(ByRef $oDoc[, $iHeight = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iHeight             - [optional] (1-25) Default is Null. The number of lines to display in the formula bar.
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current Formula Bar Height as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iHeight not an Integer, less than 1 or greater than 25.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Formula Bar height.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iHeight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version less than 7.4.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOCalc_DocColumnsRowsFreeze, _LOCalc_DocWindowVisibleRange, _LOCalc_DocWindowSplit, _LOCalc_DocWindowFirstColumn, _LOCalc_DocWindowFirstRow, _LOCalc_DocViewWindowSettings
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocFormulaBarHeight(ByRef $oDoc, $iHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurHeight

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.4) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	If __LO_VarsAreNull($iHeight) Then
		$iCurHeight = $oDoc.CurrentController.FormulaBarHeight()
		If Not IsInt($iCurHeight) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurHeight)
	EndIf

	If Not __LO_IntIsBetween($iHeight, 1, 25) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.FormulaBarHeight = $iHeight
	$iError = ($oDoc.CurrentController.FormulaBarHeight() = $iHeight) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocFormulaBarHeight

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocGetName
; Description ...: Retrieve the document's name.
; Syntax ........: _LOCalc_DocGetName(ByRef $oDoc[, $bReturnFull = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bReturnFull         - [optional] Default is False. If True, the full window title is returned, such as is used by AutoIt window related functions.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Success. Returning the document's Name as a String.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bReturnFull not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Document's name. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bReturnFull is True, the return value will be like: "<Calc Doc name>.<extension> — LibreOffice Calc" e.g. "Testing.ods — LibreOffice Calc".
;                  Else the return value will be like: "<Calc Doc name>.<extension>", e.g. "Testing.ods"
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocGetName(ByRef $oDoc, $bReturnFull = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnFull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bReturnFull Then
		$sName = $oDoc.CurrentController.Frame.Title()
		If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else
		$sName = $oDoc.Title()
		If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $sName)
EndFunc   ;==>_LOCalc_DocGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocGetPath
; Description ...: Returns a Document's current save path.
; Syntax ........: _LOCalc_DocGetPath(ByRef $oDoc[, $bReturnLibreURL = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bReturnLibreURL     - [optional] Default is False. If True, returns a path in LibreOffice URL format, else False returns a regular Windows path.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Success. Returning the document's save path as a String.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bReturnLibreURL not a Boolean.
;                  @Error: 1, @Extended: 3 = Document has no save path.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error converting LibreOffice URL to Computer path format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LO_PathConvert, _LOCalc_DocHasPath
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocGetPath(ByRef $oDoc, $bReturnLibreURL = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sPath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnLibreURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.hasLocation() Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$sPath = $oDoc.URL()

	If Not $bReturnLibreURL Then
		$sPath = _LO_PathConvert($sPath, $LO_PATHCONV_PCPATH_RETURN)
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $sPath)
EndFunc   ;==>_LOCalc_DocGetPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocHasPath
; Description ...: Returns whether a document has been saved to a location already or not.
; Syntax ........: _LOCalc_DocHasPath(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the document has a save location. Else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query Document whether it has a path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocClose, _LOCalc_DocGetPath, _LOCalc_DocSaveAs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocHasPath(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bHasPath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bHasPath = $oDoc.hasLocation()
	If Not IsBool($bHasPath) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bHasPath)
EndFunc   ;==>_LOCalc_DocHasPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocIsActive
; Description ...: Tests if called document is the active document of other LibreOffice windows.
; Syntax ........: _LOCalc_DocIsActive(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if document is the currently active LibreOffice window. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query Document whether it is active.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This does NOT test if the document is the current active window in Windows, it only tests if the document is the current active document among other LibreOffice documents.
; Related .......: _LOCalc_DocToFront
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocIsActive(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsActive

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsActive = $oDoc.CurrentController.Frame.isActive()
	If Not IsBool($bIsActive) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsActive)
EndFunc   ;==>_LOCalc_DocIsActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocIsModified
; Description ...: Set or Retrieve the document's modified status.
; Syntax ........: _LOCalc_DocIsModified(ByRef $oDoc[, $bModified = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bModified           - [optional] Default is Null. If True, sets the Document's modified status to True.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Document modified status was successfully set.
;                  @Error: 0, @Extended: 1, Return: Boolean = Success. Returning True if the document has been modified since last being saved, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bModified not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query Document whether it has been modified.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bModified
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocIsReadOnly, _LOCalc_DocSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocIsModified(ByRef $oDoc, $bModified = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsMod
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bModified) Then
		$bIsMod = $oDoc.isModified()
		If Not IsBool($bIsMod) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $bIsMod)
	EndIf

	If Not IsBool($bModified) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.Modified = $bModified
	$iError = ($oDoc.isModified() = $bModified) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocIsReadOnly
; Description ...: Tests whether a document is opened in Read Only mode.
; Syntax ........: _LOCalc_DocIsReadOnly(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True is document is currently Read Only, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query whether Document is Read-Only.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only documents that have been saved to a location, will ever be "ReadOnly".
; Related .......: _LOCalc_DocIsModified, _LOCalc_DocClose, _LOCalc_DocOpen
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocIsReadOnly(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsReadOnly

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsReadOnly = $oDoc.isReadOnly()
	If Not IsBool($bIsReadOnly) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsReadOnly)
EndFunc   ;==>_LOCalc_DocIsReadOnly

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocMaximize
; Description ...: Maximize or restore a document.
; Syntax ........: _LOCalc_DocMaximize(ByRef $oDoc[, $bMaximize = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bMaximize           - [optional] Default is Null. If True, document window is maximized, else if False, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Document was successfully maximized.
;                  @Error: 0, @Extended: 1, Return: Boolean = Success. $bMaximize called with Null, returning boolean indicating if Document is currently maximized (True) or not (False).
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bMaximize not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query whether Document is Maximized.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bMaximize
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOCalc_DocMinimize, _LOCalc_DocPosAndSize
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocMaximize(ByRef $oDoc, $bMaximize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsMax
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bMaximize) Then
		$bIsMax = $oDoc.CurrentController.Frame.ContainerWindow.IsMaximized()
		If Not IsBool($bIsMax) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $bIsMax)
	EndIf

	If Not IsBool($bMaximize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.IsMaximized = $bMaximize
	$iError = ($oDoc.CurrentController.Frame.ContainerWindow.IsMaximized() = $bMaximize) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocMaximize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocMinimize
; Description ...: Minimize or restore a document.
; Syntax ........: _LOCalc_DocMinimize(ByRef $oDoc[, $bMinimize = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bMinimize           - [optional] Default is Null. If True, document window is minimized, else if False, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Document was successfully minimized.
;                  @Error: 0, @Extended: 1, Return: Boolean = Success. $bMinimize called with Null, returning boolean indicating if Document is currently minimized (True) or not (False).
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bMinimize not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query whether Document is Minimized.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bMinimize
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOCalc_DocMaximize, _LOCalc_DocToFront, _LOCalc_DocPosAndSize
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocMinimize(ByRef $oDoc, $bMinimize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsMin
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bMinimize) Then
		$bIsMin = $oDoc.CurrentController.Frame.ContainerWindow.IsMinimized()
		If Not IsBool($bIsMin) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $bIsMin)
	EndIf

	If Not IsBool($bMinimize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.IsMinimized = $bMinimize
	$iError = ($oDoc.CurrentController.Frame.ContainerWindow.IsMinimized() = $bMinimize) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocMinimize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocOpen
; Description ...: Open an existing Calc Document, returning its object identifier.
; Syntax ........: _LOCalc_DocOpen($sFilePath[, $bConnectIfOpen = True[, $bHidden = Null[, $bReadOnly = Null[, $sPassword = Null[, $bLoadAsTemplate = Null[, $sFilterName = Null]]]]]])
; Parameters ....: $sFilePath           - Full path and filename of the file to be opened.
;                  $bConnectIfOpen      - [optional] Default is True. If True, connect to the requested document if it is already open. See remarks.
;                  $bHidden             - [optional] Default is Null. If True, opens the document invisibly.
;                  $bReadOnly           - [optional] Default is Null. If True, opens the document as read-only.
;                  $sPassword           - [optional] Default is Null. The password that was used to read-protect the document, if any.
;                  $bLoadAsTemplate     - [optional] Default is Null. If True, opens the document as a Template, i.e. an untitled copy of the specified document is made instead of modifying the original document.
;                  $sFilterName         - [optional] Default is Null. Name of a LibreOffice filter to use to load the specified document. LibreOffice automatically selects which to use by default.
; Return values .: Success: Object.
;                  @Error: 0, @Extended: 1, Return: Object = Successfully connected to requested Document without requested parameters. Returning Document's Object.
;                  @Error: 0, @Extended: 2, Return: Object = Successfully opened requested Document with requested parameters. Returning Document's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $sFilePath not string, or file not found.
;                  @Error: 1, @Extended: 2 = Error converting file path to URL path.
;                  @Error: 1, @Extended: 3 = $bConnectIfOpen not a Boolean.
;                  @Error: 1, @Extended: 4 = $bHidden not a Boolean.
;                  @Error: 1, @Extended: 5 = $bReadOnly not a Boolean.
;                  @Error: 1, @Extended: 6 = $sPassword not a string.
;                  @Error: 1, @Extended: 7 = $bLoadAsTemplate not a Boolean.
;                  @Error: 1, @Extended: 8 = $sFilterName not a string.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create ServiceManager Object
;                  @Error: 2, @Extended: 2 = Failed to create Desktop Object
;                  @Error: 2, @Extended: 3 = Failed opening or connecting to document.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHidden
;                  |                               2 = Error setting $bReadOnly
;                  |                               4 = Error setting $sPassword
;                  |                               8 = Error setting $bLoadAsTemplate
;                  |                               16 = Error setting $sFilterName
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Any parameters (Hidden, template etc.,) will not be applied when connecting to a document.
; Related .......: _LOCalc_DocCreate, _LOCalc_DocClose, _LOCalc_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocOpen($sFilePath, $bConnectIfOpen = True, $bHidden = Null, $bReadOnly = Null, $sPassword = Null, $bLoadAsTemplate = Null, $sFilterName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $iError = 0
	Local $oDoc, $oServiceManager, $oDesktop
	Local $aoProperties[0]
	Local $vProperty
	Local $sFileURL

	If Not IsString($sFilePath) Or Not FileExists($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sFileURL = _LO_PathConvert($sFilePath, $LO_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bConnectIfOpen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not __LO_VarsAreNull($bHidden, $bReadOnly, $sPassword, $bLoadAsTemplate, $sFilterName) Then
		If ($bHidden <> Null) Then
			If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$vProperty = __LO_SetPropertyValue("Hidden", $bHidden)
			If @error Then $iError = BitOR($iError, 1)
			If Not BitAND($iError, 1) Then __LO_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($bReadOnly <> Null) Then
			If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$vProperty = __LO_SetPropertyValue("ReadOnly", $bReadOnly)
			If @error Then $iError = BitOR($iError, 2)
			If Not BitAND($iError, 2) Then __LO_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($sPassword <> Null) Then
			If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$vProperty = __LO_SetPropertyValue("Password", $sPassword)
			If @error Then $iError = BitOR($iError, 4)
			If Not BitAND($iError, 4) Then __LO_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($bLoadAsTemplate <> Null) Then
			If Not IsBool($bLoadAsTemplate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$vProperty = __LO_SetPropertyValue("AsTemplate", $bLoadAsTemplate)
			If @error Then $iError = BitOR($iError, 8)
			If Not BitAND($iError, 8) Then __LO_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($sFilterName <> Null) Then
			If Not IsString($sFilterName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$vProperty = __LO_SetPropertyValue("FilterName", $sFilterName)
			If @error Then $iError = BitOR($iError, 16)
			If Not BitAND($iError, 16) Then __LO_AddTo1DArray($aoProperties, $vProperty)
		EndIf
	EndIf

	If $bConnectIfOpen Then $oDoc = _LOCalc_DocConnect($LO_DOC_CONNECT_MODE_SEARCH_PATH, $sFilePath, True)
	If IsObj($oDoc) Then Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))

	$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOCalc_DocOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocPosAndSize
; Description ...: Reposition and resize a document window.
; Syntax ........: _LOCalc_DocPosAndSize(ByRef $oDoc[, $iX = Null[, $iY = Null[, $iWidth = Null[, $iHeight = Null]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iX                  - [optional] Default is Null. The X coordinate of the window.
;                  $iY                  - [optional] Default is Null. The Y coordinate of the window.
;                  $iWidth              - [optional] Default is Null. The width of the window, in pixels(?).
;                  $iHeight             - [optional] Default is Null. The height of the window, in pixels(?).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iX not an Integer.
;                  @Error: 1, @Extended: 3 = $iY not an Integer.
;                  @Error: 1, @Extended: 4 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 5 = $iHeight not an Integer.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Position and Size Structure Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Position and Size Structure Object for error checking.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iX
;                  |                               2 = Error setting $iY
;                  |                               4 = Error setting $iWidth
;                  |                               8 = Error setting $iHeight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: X & Y, on my computer at least, seem to go no lower than 8(X) and 30(Y), if you enter lower than this, it will cause a "property setting Error".
;                  If you want more accurate functionality, use the "WinMove" AutoIt function.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOCalc_DocMaximize, _LOCalc_DocMinimize, _LOCalc_DocToFront, _LOCalc_DocZoom
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocPosAndSize(ByRef $oDoc, $iX = Null, $iY = Null, $iWidth = Null, $iHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tWindowSize
	Local Const $iPosSize = 15 ; adjust both size and position.
	Local $iError = 0
	Local $aiWinPosSize[4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tWindowSize = $oDoc.CurrentController.Frame.ContainerWindow.getPosSize()
	If Not IsObj($tWindowSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iX, $iY, $iWidth, $iHeight) Then
		__LO_ArrayFill($aiWinPosSize, $tWindowSize.X(), $tWindowSize.Y(), $tWindowSize.Width(), $tWindowSize.Height())

		Return SetError($__LO_STATUS_SUCCESS, 2, $aiWinPosSize)
	EndIf

	If ($iX <> Null) Then
		If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tWindowSize.X = $iX
	EndIf

	If ($iY <> Null) Then
		If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tWindowSize.Y = $iY
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tWindowSize.Width = $iWidth
	EndIf

	If ($iHeight <> Null) Then
		If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tWindowSize.Height = $iHeight
	EndIf

	$oDoc.CurrentController.Frame.ContainerWindow.setPosSize($tWindowSize.X, $tWindowSize.Y, $tWindowSize.Width, $tWindowSize.Height, $iPosSize)

	$tWindowSize = $oDoc.CurrentController.Frame.ContainerWindow.getPosSize()
	If Not IsObj($tWindowSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iError = (__LO_VarsAreNull($iX)) ? ($iError) : (($tWindowSize.X() = $iX) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($iY)) ? ($iError) : (($tWindowSize.Y() = $iY) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iWidth)) ? ($iError) : (($tWindowSize.Width() = $iWidth) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iHeight)) ? ($iError) : (($tWindowSize.Height() = $iHeight) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocPosAndSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocPrint
; Description ...: Print a document using the specified settings.
; Syntax ........: _LOCalc_DocPrint(ByRef $oDoc[, $iCopies = 1[, $bCollate = True[, $vPages = "ALL"[, $bWait = True[, $iDuplexMode = $LOC_PRINT_DUPLEX_OFF[, $sPrinter = ""[, $sFilePathName = ""]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iCopies             - [optional] Default is 1. Specifies the number of copies to print.
;                  $bCollate            - [optional] Default is True. Advises the printer to collate the pages of the copies.
;                  $vPages              - [optional] Default is "ALL". Specifies which pages to print. See remarks.
;                  $bWait               - [optional] Default is True. If True, the corresponding print request will be executed synchronous. Default is to use synchronous print mode.
;                  $iDuplexMode         - [optional] (0-3) Default is $LOC_PRINT_DUPLEX_OFF. Determines the duplex mode for the print job. See Constants, $LOC_PRINT_DUPLEX_* as defined in LibreOfficeCalc_Constants.au3.
;                  $sPrinter            - [optional] Default is "". Printer name. If left blank, or if printer name is not found, default printer is used.
;                  $sFilePathName       - [optional] Default is "". Specifies the name of a file to print to. Creates a .prn file at the given Path. Must include the desired path destination with file name.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success Document was successfully printed.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iCopies not a Integer.
;                  @Error: 1, @Extended: 3 = $bCollate not a Boolean.
;                  @Error: 1, @Extended: 4 = $vPages not an Integer or String.
;                  @Error: 1, @Extended: 5 = $vPages contains invalid characters, a-z, or a period(.).
;                  @Error: 1, @Extended: 6 = $bWait not a Boolean.
;                  @Error: 1, @Extended: 7 = $iDuplexMode not an Integer, less than 0 or greater than 3. See Constants, $LOC_PRINT_DUPLEX_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error: 1, @Extended: 8 = $sPrinter not a String.
;                  @Error: 1, @Extended: 9 = $sFilePathName not a String.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "Printer Name" property.
;                  @Error: 2, @Extended: 2 = Error creating "Copies" property.
;                  @Error: 2, @Extended: 3 = Error creating "Collate" property.
;                  @Error: 2, @Extended: 4 = Error creating "Wait" property.
;                  @Error: 2, @Extended: 5 = Error creating "DuplexMode" property.
;                  @Error: 2, @Extended: 6 = Error creating "Pages" property.
;                  @Error: 2, @Extended: 7 = Error creating "PrintToFile" property.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error converting PrintToFile Path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Setting $bWait to True is highly recommended. Otherwise following actions (as e.g. closing the Document) can fail.
;                  Based on OOoCalc UDF Print function by GMK.
;                  $vPages range can be called as entered in the user interface, as follows: "1-4,10" to print the pages 1 to 4 and 10. Default is "ALL". Must be in String format to accept more than just a single page number. e.g. This will work: "1-6,12,27" This will not 1-6,12,27. This will work: "7", This will also: 7.
;                  To set the output paper size, you would have to modify the Page Style used for the sheet.
; Related .......: _LOCalc_SheetPrintColumnsRepeat, _LOCalc_SheetPrintRowsRepeat
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocPrint(ByRef $oDoc, $iCopies = 1, $bCollate = True, $vPages = "ALL", $bWait = True, $iDuplexMode = $LOC_PRINT_DUPLEX_OFF, $sPrinter = "", $sFilePathName = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__STR_STRIPLEADING = 1, $__STR_STRIPTRAILING = 2, $__STR_STRIPALL = 8
	Local $avPrintOpt[4], $asSetPrinterOpt[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iCopies) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bCollate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($vPages) And Not IsString($vPages) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$vPages = (IsString($vPages)) ? (StringStripWS($vPages, $__STR_STRIPALL)) : ($vPages)
	If IsString($vPages) And Not ($vPages = "ALL") Then
		If StringRegExp($vPages, "[[:alpha:]]|[\.]") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	EndIf
	If Not IsBool($bWait) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not __LO_IntIsBetween($iDuplexMode, $LOC_PRINT_DUPLEX_OFF, $LOC_PRINT_DUPLEX_SHORT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not IsString($sPrinter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	$sPrinter = StringStripWS(StringStripWS($sPrinter, $__STR_STRIPTRAILING), $__STR_STRIPLEADING)
	If Not IsString($sFilePathName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	$sFilePathName = StringStripWS(StringStripWS($sFilePathName, $__STR_STRIPTRAILING), $__STR_STRIPLEADING)
	If $sPrinter <> "" Then
		$asSetPrinterOpt[0] = __LO_SetPropertyValue("Name", $sPrinter)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDoc.setPrinter($asSetPrinterOpt)
	EndIf
	$avPrintOpt[0] = __LO_SetPropertyValue("CopyCount", $iCopies)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$avPrintOpt[1] = __LO_SetPropertyValue("Collate", $bCollate)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	$avPrintOpt[2] = __LO_SetPropertyValue("Wait", $bWait)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$avPrintOpt[3] = __LO_SetPropertyValue("DuplexMode", $iDuplexMode)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)

	If $vPages <> "ALL" Then
		ReDim $avPrintOpt[UBound($avPrintOpt) + 1]
		$avPrintOpt[UBound($avPrintOpt) - 1] = __LO_SetPropertyValue("Pages", $vPages)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 6, 0)
	EndIf
	If $sFilePathName <> "" Then
		$sFilePathName = $sFilePathName & ".prn"
		$sFilePathName = _LO_PathConvert($sFilePathName, $LO_PATHCONV_OFFICE_RETURN)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		ReDim $avPrintOpt[UBound($avPrintOpt) + 1]
		$avPrintOpt[UBound($avPrintOpt) - 1] = __LO_SetPropertyValue("FileName", $sFilePathName)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 7, 0)
	EndIf
	$oDoc.print($avPrintOpt)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocPrint

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocRedo
; Description ...: Perform one Redo action for a document.
; Syntax ........: _LOCalc_DocRedo(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Successfully performed a redo action.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Document does not have a redo action to perform.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocRedoIsPossible, _LOCalc_DocRedoGetAllActionTitles, _LOCalc_DocRedoCurActionTitle, _LOCalc_DocUndo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocRedo(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isRedoPossible()) Then
		$oDoc.UndoManager.Redo()

		Return SetError($__LO_STATUS_SUCCESS, 1, 0)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOCalc_DocRedo

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocRedoClear
; Description ...: Clear all Redo Actions in the Redo Action List.
; Syntax ........: _LOCalc_DocRedoClear(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully cleared all Redo Actions.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This will silently fail if there are any _LOCalc_DocUndoActionBegin still active.
; Related .......: _LOCalc_DocUndoReset, _LOCalc_DocUndoClear
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocRedoClear(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.clearRedo()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocRedoClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocRedoCurActionTitle
; Description ...: Retrieve the current Redo action Title.
; Syntax ........: _LOCalc_DocRedoCurActionTitle(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Returning the current available redo action title as a String. Will be an empty String if no action is available.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Redo Action.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocRedo, _LOCalc_DocRedoGetAllActionTitles, _LOCalc_DocRedoIsPossible, _LOCalc_DocUndoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocRedoCurActionTitle(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sRedoAction

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sRedoAction = $oDoc.UndoManager.getCurrentRedoActionTitle()
	If ($sRedoAction = Null) Then $sRedoAction = ""
	If Not IsString($sRedoAction) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sRedoAction)
EndFunc   ;==>_LOCalc_DocRedoCurActionTitle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocRedoGetAllActionTitles
; Description ...: Retrieve all available Redo action Titles.
; Syntax ........: _LOCalc_DocRedoGetAllActionTitles(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Array.
;                  @Error: 0, @Extended: ?, Return: Array = Returning all available redo action Titles in an array of Strings. @Extended set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve an array of Redo action titles.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocRedo, _LOCalc_DocRedoCurActionTitle, _LOCalc_DocUndoGetAllActionTitles
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocRedoGetAllActionTitles(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asTitles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asTitles = $oDoc.UndoManager.getAllRedoActionTitles()
	If Not IsArray($asTitles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asTitles), $asTitles)
EndFunc   ;==>_LOCalc_DocRedoGetAllActionTitles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocRedoIsPossible
; Description ...: Test whether a Redo action is available to perform for a document.
; Syntax ........: _LOCalc_DocRedoIsPossible(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = If the document has a redo action to perform, True is returned, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query whether a Redo is possible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocRedo, _LOCalc_DocRedoCurActionTitle, _LOCalc_DocRedoGetAllActionTitles, _LOCalc_DocUndoIsPossible
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocRedoIsPossible(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsRedoPoss

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsRedoPoss = $oDoc.UndoManager.isRedoPossible()
	If Not IsBool($bIsRedoPoss) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $bIsRedoPoss)
EndFunc   ;==>_LOCalc_DocRedoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocSave
; Description ...: Save any changes made to a Document.
; Syntax ........: _LOCalc_DocSave(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Document Successfully saved.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Document is Read Only or Document has no save location, try SaveAs.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocExport, _LOCalc_DocSaveAs, _LOCalc_DocIsModified, _LOCalc_DocIsReadOnly, _LOCalc_DocHasPath
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocSave(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDoc.hasLocation Or $oDoc.isReadOnly Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oDoc.store()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocSaveAs
; Description ...: Save a Document with the specified file name to the path specified with any parameters called.
; Syntax ........: _LOCalc_DocSaveAs(ByRef $oDoc, $sFilePath[, $sFilterName = ""[, $bOverwrite = Null[, $sPassword = Null]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sFilePath           - Full path to save the document to, including Filename and extension.
;                  $sFilterName         - [optional] Default is "". The filter name. Calling "" (blank string), means the filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .ods extension is used instead, with the filter name of "calc8".
;                  $bOverwrite          - [optional] Default is Null. If True, the existing file will be overwritten.
;                  $sPassword           - [optional] Default is Null. Sets a password for the document. (Not all file formats can have a Password set). Null or "" (blank string) = No Password.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Successfully Saved the document. Returning document save path.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sFilePath not a String.
;                  @Error: 1, @Extended: 3 = $sFilterName not a String.
;                  @Error: 1, @Extended: 4 = $bOverwrite not a Boolean.
;                  @Error: 1, @Extended: 5 = $sPassword not a String.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating FilterName Property
;                  @Error: 2, @Extended: 2 = Error creating Overwrite Property
;                  @Error: 2, @Extended: 3 = Error creating Password Property
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error Converting Path to/from L.O. URL
;                  @Error: 3, @Extended: 2 = Error retrieving FilterName.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Alters original save path (if there was one) to the new path.
; Related .......: _LOCalc_DocExport, _LOCalc_DocSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocSaveAs(ByRef $oDoc, $sFilePath, $sFilterName = "", $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aProperties[1]
	Local $sSavePath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFilterName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$sFilePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFilterName = "") Or ($sFilterName = " ") Then $sFilterName = __LOCalc_FilterNameGet($sFilePath)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$aProperties[0] = __LO_SetPropertyValue("FilterName", $sFilterName)
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	EndIf

	If $sPassword <> Null Then
		If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)
	EndIf

	$oDoc.storeAsURL($sFilePath, $aProperties)

	$sSavePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOCalc_DocSaveAs

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocSelectionCopy
; Description ...: "Copies" data selected by the Cursor, returning an Object for use in inserting later.
; Syntax ........: _LOCalc_DocSelectionCopy(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Data was successfully copied, returning an Object for use in _LOCalc_DocSelectionPaste.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to Copy Selected Data, make sure Data is selected.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Data you desire to be copied must be selected first, see _LOCalc_DocSelectionSet, _LOCalc_DocSelectionSetMulti.
;                  This function works essentially the same as Copy/ Ctrl+C, except it doesn't use your clipboard.
;                  The Object returned is used in _LOCalc_DocSelectionPaste to insert the data again.
;                  Data copied can be inserted into the same or another document.
; Related .......: _LOCalc_DocSelectionPaste, _LOCalc_DocSelectionSet, _LOCalc_DocSelectionSetMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocSelectionCopy(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oObj = $oDoc.CurrentController.getTransferable() ; Copy
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oObj)
EndFunc   ;==>_LOCalc_DocSelectionCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocSelectionGet
; Description ...: Retrieve the current user selection(s).
; Syntax ........: _LOCalc_DocSelectionGet(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Object or Array
;                  @Error: 0, @Extended: 0, Return: Object = Success. Single cell selected or cursor is editing a cell, returning Cell Object.
;                  @Error: 0, @Extended: 1, Return: Object = Success. Cell Range selected, returning Cell Range Object.
;                  @Error: 0, @Extended: ?, Return: Array = Success. Multiple Cells or Cell Ranges selected, returning array of Cell Range Objects.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current selection.
;                  @Error: 3, @Extended: 2 = Failed to retrieve count of multiple selections.
;                  @Error: 3, @Extended: 3 = Failed to determine selection type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the user has nothing selected, or is typing in a cell, the return will still be the single cell Object.
; Related .......: _LOCalc_DocSelectionCopy, _LOCalc_DocSelectionSet, _LOCalc_DocSelectionSetMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocSelectionGet(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSelection
	Local $aoSelections[0]
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSelection = $oDoc.getCurrentSelection()
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Select
		Case $oSelection.supportsService("com.sun.star.sheet.SheetCell") ; Single Cell is selected.

			Return SetError($__LO_STATUS_SUCCESS, 0, $oSelection)

		Case $oSelection.supportsService("com.sun.star.sheet.SheetCellRange") ; Single Range is selected.

			Return SetError($__LO_STATUS_SUCCESS, 1, $oSelection)

		Case $oSelection.supportsService("com.sun.star.sheet.SheetCellRanges") ; Multiple Ranges are selected.
			$iCount = $oSelection.getCount()
			If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			ReDim $aoSelections[$iCount]

			For $i = 0 To $iCount - 1
				$aoSelections[$i] = $oSelection.getByIndex($i)
			Next

			Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoSelections)
	EndSelect

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
EndFunc   ;==>_LOCalc_DocSelectionGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocSelectionPaste
; Description ...: Inserts a ParObjCopy Object at the current ViewCursor location.
; Syntax ........: _LOCalc_DocSelectionPaste(ByRef $oDoc, ByRef $oData)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oData               - A Object returned from _LOCalc_DocSelectionCopy to insert.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Data was successfully inserted at the currently selected cell.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oData not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The data will be pasted into the document, beginning at the currently selected cell.
; Related .......: _LOCalc_DocSelectionCopy, _LOCalc_DocSelectionSet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocSelectionPaste(ByRef $oDoc, ByRef $oData)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oData) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.insertTransferable($oData)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocSelectionPaste

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocSelectionSet
; Description ...: Set the current selection for the Document.
; Syntax ........: _LOCalc_DocSelectionSet(ByRef $oDoc, ByRef $oObj)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oObj                - A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetActive function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Object called in $oObj successfully selected.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oObj not an Object.
;                  @Error: 1, @Extended: 3 = Object called in $oObj not a Cell Object and not a Cell Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocSelectionCopy, _LOCalc_DocSelectionGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocSelectionSet(ByRef $oDoc, ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oObj.supportsService("com.sun.star.sheet.SheetCell") And Not _
			$oObj.supportsService("com.sun.star.sheet.SheetCellRange") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oDoc.CurrentController.Select($oObj)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocSelectionSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocSelectionSetMulti
; Description ...: Select multiple Ranges in a Document.
; Syntax ........: _LOCalc_DocSelectionSetMulti(ByRef $oDoc, ByRef $aoRange)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $aoRange             - An array of Cell or Cell Range objects returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetActive function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Objects were successfully selected.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an object.
;                  @Error: 1, @Extended: 2 = $aoRange not an Array.
;                  @Error: 1, @Extended: 3 = Array called in $aoRange does not contain an Object. Returning problem element index.
;                  @Error: 1, @Extended: 4 = Array called in $aoRange contains an Object that is not a Cell Object and not a Cell Range. Returning problem element index.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a "com.sun.star.sheet.SheetCellRanges" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Range Address from Object located in array called in $aoRange. Returning problem element index.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocSelectionCopy,_LOCalc_DocSelectionSet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocSelectionSetMulti(ByRef $oDoc, ByRef $aoRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRanges
	Local $aoRangeAddr[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($aoRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	ReDim $aoRangeAddr[UBound($aoRange)]

	For $i = 0 To UBound($aoRange) - 1
		If Not IsObj($aoRange[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, $i)
		If Not $aoRange[$i].supportsService("com.sun.star.sheet.SheetCell") And Not _
				$aoRange[$i].supportsService("com.sun.star.sheet.SheetCellRange") Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
		$aoRangeAddr[$i] = $aoRange[$i].RangeAddress()
		If Not IsObj($aoRangeAddr[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, $i)

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$oRanges = $oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
	If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oRanges.addRangeAddresses($aoRangeAddr, True)

	$oDoc.CurrentController.Select($oRanges)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocSelectionSetMulti

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocToFront
; Description ...: Bring the called document to the front of the other windows.
; Syntax ........: _LOCalc_DocToFront(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Window was successfully brought to the front of the open windows.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If minimized, the document is restored and brought to the front of the visible pages. Generally only brings the document to the front of other LibreOffice windows.
; Related .......: _LOCalc_DocIsActive, _LOCalc_DocMinimize
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocToFront(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.toFront()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocToFront

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndo
; Description ...: Perform one Undo action for a document.
; Syntax ........: _LOCalc_DocUndo(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Successfully performed an undo action.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Document does not have an undo action to perform.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocUndoIsPossible, _LOCalc_DocUndoGetAllActionTitles, _LOCalc_DocUndoCurActionTitle, _LOCalc_DocRedo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndo(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isUndoPossible()) Then
		$oDoc.UndoManager.Undo()

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOCalc_DocUndo

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndoActionBegin
; Description ...: Begin an Undo Action group.
; Syntax ........: _LOCalc_DocUndoActionBegin(ByRef $oDoc[, $sName = "AU3LO-Automation"])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sName               - [optional] Default is "AU3LO-Automation". The name of the Undo Action to display in the UI when completed.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully began an Undo Action group recording.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This begins an Undo Action Group, any functions and actions done after this function is called will be grouped together, and if undone, all actions will be undone together at once.
;                  _LOCalc_DocUndoActionEnd must be called after this function before this undo group will become available in the Undo Action list.
;                  _LOCalc_DocUndoActionBegin can be nested, e.g. call this function multiple times without ending the first undo action, but only the last group that is ended with _LOCalc_DocUndoActionEnd will appear.
; Related .......: _LOCalc_DocUndoActionEnd, _LOCalc_DocUndoReset
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndoActionBegin(ByRef $oDoc, $sName = "AU3LO-Automation")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.UndoManager.enterUndoContext($sName)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocUndoActionBegin

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndoActionEnd
; Description ...: End the last started Undo Action Group.
; Syntax ........: _LOCalc_DocUndoActionEnd(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully ended the last Undo Action group recording.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This stops the grouping of actions into the last created Undo Action Group.
; Related .......: _LOCalc_DocUndoActionBegin, _LOCalc_DocUndoReset
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndoActionEnd(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.leaveUndoContext()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocUndoActionEnd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndoClear
; Description ...: Clear all Undo and Redo Actions in the Undo/Redo Action List.
; Syntax ........: _LOCalc_DocUndoClear(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully cleared all Undo and Redo Actions.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This will silently fail if there are any _LOCalc_DocUndoActionBegin still active.
; Related .......: _LOCalc_DocUndoReset, _LOCalc_DocRedoClear
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndoClear(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.clear()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocUndoClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndoCurActionTitle
; Description ...: Retrieve the current Undo action Title.
; Syntax ........: _LOCalc_DocUndoCurActionTitle(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Returning the current available Undo action title as a String. Will be an empty String if no action is available.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Undo Action.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocUndo, _LOCalc_DocUndoGetAllActionTitles, _LOCalc_DocUndoIsPossible, _LOCalc_DocRedoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndoCurActionTitle(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sUndoAction
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sUndoAction = $oDoc.UndoManager.getCurrentUndoActionTitle()
	If ($sUndoAction = Null) Then $sUndoAction = ""
	If Not IsString($sUndoAction) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sUndoAction)
EndFunc   ;==>_LOCalc_DocUndoCurActionTitle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndoGetAllActionTitles
; Description ...: Retrieve all available Undo action Titles.
; Syntax ........: _LOCalc_DocUndoGetAllActionTitles(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Array.
;                  @Error: 0, @Extended: ?, Return: Array = Returning all available undo action Titles in an array of Strings. @Extended set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve an array of Undo action titles.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocUndo, _LOCalc_DocUndoCurActionTitle, _LOCalc_DocRedoGetAllActionTitles
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndoGetAllActionTitles(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asTitles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asTitles = $oDoc.UndoManager.getAllUndoActionTitles()
	If Not IsArray($asTitles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asTitles), $asTitles)
EndFunc   ;==>_LOCalc_DocUndoGetAllActionTitles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndoIsPossible
; Description ...: Test whether a Undo action is available to perform for a document.
; Syntax ........: _LOCalc_DocUndoIsPossible(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = If the document has an undo action to perform, True is returned, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query whether an Undo is possible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocUndo, _LOCalc_DocRedoIsPossible
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndoIsPossible(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsUndoPoss

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsUndoPoss = $oDoc.UndoManager.isUndoPossible()
	If Not IsBool($bIsUndoPoss) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $bIsUndoPoss)
EndFunc   ;==>_LOCalc_DocUndoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndoReset
; Description ...: Reset the UndoManager.
; Syntax ........: _LOCalc_DocUndoReset(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully reset the undo manager.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling this function does the following: remove all locks from the undo manager; closes all open undo group actions, clears all undo actions, clears all redo actions.
; Related .......: _LOCalc_DocUndoClear, _LOCalc_DocRedoClear
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndoReset(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.reset()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocUndoReset

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocViewDisplaySettings
; Description ...: Set or Retrieve the current Document View Display settings.
; Syntax ........: _LOCalc_DocViewDisplaySettings(ByRef $oDoc[, $bFormulas = Null[, $bZeroValues = Null[, $bComments = Null[, $bPageBreaks = Null[, $bHelpLines = Null[, $bValueHighlight = Null[, $bAnchors = Null[, $bGrid = Null[, $iGridColor = Null]]]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bFormulas           - [optional] Default is Null. If True, Formulas, rather than results, are displayed in the cells.
;                  $bZeroValues         - [optional] Default is Null. If True, numbers with the value of 0 are shown.
;                  $bComments           - [optional] Default is Null. If True, a small rectangle in the top right corner of the cell indicates that a comment exists.
;                  $bPageBreaks         - [optional] Default is Null. If True, Page Breaks are displayed for a print area.
;                  $bHelpLines          - [optional] Default is Null. If True, help lines are displayed while moving graphics, drawings, etc.
;                  $bValueHighlight     - [optional] Default is Null. If True, Cell contents are displayed in different colors, depending on the content type of the cell.
;                  $bAnchors            - [optional] Default is Null. If True, the Anchor icon is displayed when a graphic or other object is selected.
;                  $bGrid               - [optional] Default is Null. If True, Gridlines are displayed.
;                  $iGridColor          - [optional] (0-16777215) Default is Null. The Grid line color, as a RGB Color Integer. Can be one of the constants $LO_COLOR_* as defined in LibreOffice_Constants.au3 or a custom value.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 9 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bFormulas not a Boolean.
;                  @Error: 1, @Extended: 3 = $bZeroValues not a Boolean.
;                  @Error: 1, @Extended: 4 = $bComments not a Boolean.
;                  @Error: 1, @Extended: 5 = $bPageBreaks not a Boolean.
;                  @Error: 1, @Extended: 6 = $bHelpLines not a Boolean.
;                  @Error: 1, @Extended: 7 = $bValueHighlight not a Boolean.
;                  @Error: 1, @Extended: 8 = $bAnchors not a Boolean.
;                  @Error: 1, @Extended: 9 = $bGrid not a Boolean.
;                  @Error: 1, @Extended: 10 = $iGridColor not an Integer, less than 0 or greater than 16777215.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Current Controller Object for the Document.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFormulas
;                  |                               2 = Error setting $bZeroValues
;                  |                               4 = Error setting $bComments
;                  |                               8 = Error setting $bPageBreaks
;                  |                               16 = Error setting $bHelpLines
;                  |                               32 = Error setting $bValueHighlight
;                  |                               64 = Error setting $bAnchors
;                  |                               128 = Error setting $bGrid
;                  |                               256 = Error setting $iGridColor
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOCalc_DocViewWindowSettings, _LO_ConvertColorToLong, _LO_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocViewDisplaySettings(ByRef $oDoc, $bFormulas = Null, $bZeroValues = Null, $bComments = Null, $bPageBreaks = Null, $bHelpLines = Null, $bValueHighlight = Null, $bAnchors = Null, $bGrid = Null, $iGridColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $abView[9]
	Local $oCurCont
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oCurCont = $oDoc.CurrentController()
	If Not IsObj($oCurCont) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($bFormulas, $bZeroValues, $bComments, $bPageBreaks, $bHelpLines, $bValueHighlight, $bAnchors, $bGrid, $iGridColor) Then
		__LO_ArrayFill($abView, $oCurCont.ShowFormulas(), $oCurCont.ShowZeroValues(), $oCurCont.ShowNotes(), $oCurCont.ShowPageBreaks(), $oCurCont.ShowHelpLines(), _
				$oCurCont.IsValueHighlightingEnabled(), $oCurCont.ShowAnchor(), $oCurCont.ShowGrid(), $oCurCont.GridColor())

		Return SetError($__LO_STATUS_SUCCESS, 1, $abView)
	EndIf

	If ($bFormulas <> Null) Then
		If Not IsBool($bFormulas) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oCurCont.ShowFormulas = $bFormulas
		$iError = ($oCurCont.ShowFormulas() = $bFormulas) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bZeroValues <> Null) Then
		If Not IsBool($bZeroValues) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCurCont.ShowZeroValues = $bZeroValues
		$iError = ($oCurCont.ShowZeroValues() = $bZeroValues) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bComments <> Null) Then
		If Not IsBool($bComments) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCurCont.ShowNotes = $bComments
		$iError = ($oCurCont.ShowNotes() = $bComments) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bPageBreaks <> Null) Then
		If Not IsBool($bPageBreaks) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oCurCont.ShowPageBreaks = $bPageBreaks
		$iError = ($oCurCont.ShowPageBreaks() = $bPageBreaks) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bHelpLines <> Null) Then
		If Not IsBool($bHelpLines) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oCurCont.ShowHelpLines = $bHelpLines
		$iError = ($oCurCont.ShowHelpLines() = $bHelpLines) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bValueHighlight <> Null) Then
		If Not IsBool($bValueHighlight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oCurCont.IsValueHighlightingEnabled = $bValueHighlight
		$iError = ($oCurCont.IsValueHighlightingEnabled() = $bValueHighlight) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bAnchors <> Null) Then
		If Not IsBool($bAnchors) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oCurCont.ShowAnchor = $bAnchors
		$iError = ($oCurCont.ShowAnchor() = $bAnchors) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bGrid <> Null) Then
		If Not IsBool($bGrid) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oCurCont.ShowGrid = $bGrid
		$iError = ($oCurCont.ShowGrid() = $bGrid) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($iGridColor <> Null) Then
		If Not __LO_IntIsBetween($iGridColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oCurCont.GridColor = $iGridColor
		$iError = ($oCurCont.GridColor() = $iGridColor) ? ($iError) : (BitOR($iError, 256))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocViewDisplaySettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocViewWindowSettings
; Description ...: Set or Retrieve the current Document View Window settings.
; Syntax ........: _LOCalc_DocViewWindowSettings(ByRef $oDoc[, $bHeaders = Null[, $bVertScroll = Null[, $bHoriScroll = Null[, $bSheetTabs = Null[, $bOutlineSymbols = Null[, $bCharts = Null[, $bDrawing = Null[, $bObjects = Null]]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bHeaders            - [optional] Default is Null. If True, Column/Row headers are displayed.
;                  $bVertScroll         - [optional] Default is Null. If True, a Vertical scrollbar is displayed at the right of the document.
;                  $bHoriScroll         - [optional] Default is Null. If True, a Horizontal scrollbar is displayed at the bottom of the document.
;                  $bSheetTabs          - [optional] Default is Null. If True, Sheet Tabs selector will be displayed at the bottom of the document.
;                  $bOutlineSymbols     - [optional] Default is Null. If True, the predefined outline symbols will be displayed.
;                  $bCharts             - [optional] Default is Null. If True, Charts are visible in the document.
;                  $bDrawing            - [optional] Default is Null. If True, Drawing Objects are visible in the document.
;                  $bObjects            - [optional] Default is Null. If True, Objects/Graphics are visible in the document.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bHeaders not a Boolean.
;                  @Error: 1, @Extended: 3 = $bVertScroll not a Boolean.
;                  @Error: 1, @Extended: 4 = $bHoriScroll not a Boolean.
;                  @Error: 1, @Extended: 5 = $bSheetTabs not a Boolean.
;                  @Error: 1, @Extended: 6 = $bOutlineSymbols not a Boolean.
;                  @Error: 1, @Extended: 7 = $bCharts not a Boolean.
;                  @Error: 1, @Extended: 8 = $bDrawing not a Boolean.
;                  @Error: 1, @Extended: 9 = $bObjects not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Current Controller Object for the Document.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bHeaders
;                  |                               2 = Error setting $bVertScroll
;                  |                               4 = Error setting $bHoriScroll
;                  |                               8 = Error setting $bSheetTabs
;                  |                               16 = Error setting $bOutlineSymbols
;                  |                               32 = Error setting $bCharts
;                  |                               64 = Error setting $bDrawing
;                  |                               128 = Error setting $bObjects
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOCalc_DocFormulaBarHeight
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocViewWindowSettings(ByRef $oDoc, $bHeaders = Null, $bVertScroll = Null, $bHoriScroll = Null, $bSheetTabs = Null, $bOutlineSymbols = Null, $bCharts = Null, $bDrawing = Null, $bObjects = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $abView[8]
	Local $oCurCont
	Local Const $__LOC_ViewObjMode_SHOW = 0, $__LOC_ViewObjMode_HIDE = 1
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oCurCont = $oDoc.CurrentController()
	If Not IsObj($oCurCont) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($bHeaders, $bVertScroll, $bHoriScroll, $bSheetTabs, $bOutlineSymbols, $bCharts, $bDrawing, $bObjects) Then
		__LO_ArrayFill($abView, $oCurCont.HasColumnRowHeaders(), $oCurCont.HasVerticalScrollBar(), $oCurCont.HasHorizontalScrollBar(), $oCurCont.HasSheetTabs(), _
				$oCurCont.IsOutlineSymbolsSet(), ($oCurCont.ShowCharts() = $__LOC_ViewObjMode_SHOW) ? (True) : (False), _
				($oCurCont.ShowDrawing() = $__LOC_ViewObjMode_SHOW) ? (True) : (False), ($oCurCont.ShowObjects() = $__LOC_ViewObjMode_SHOW) ? (True) : (False))

		Return SetError($__LO_STATUS_SUCCESS, 1, $abView)
	EndIf

	If ($bHeaders <> Null) Then
		If Not IsBool($bHeaders) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oCurCont.HasColumnRowHeaders = $bHeaders
		$iError = ($oCurCont.HasColumnRowHeaders() = $bHeaders) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bVertScroll <> Null) Then
		If Not IsBool($bVertScroll) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCurCont.HasVerticalScrollBar = $bVertScroll
		$iError = ($oCurCont.HasVerticalScrollBar() = $bVertScroll) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bHoriScroll <> Null) Then
		If Not IsBool($bHoriScroll) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCurCont.HasHorizontalScrollBar = $bHoriScroll
		$iError = ($oCurCont.HasHorizontalScrollBar() = $bHoriScroll) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bSheetTabs <> Null) Then
		If Not IsBool($bSheetTabs) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oCurCont.HasSheetTabs = $bSheetTabs
		$iError = ($oCurCont.HasSheetTabs() = $bSheetTabs) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bOutlineSymbols <> Null) Then
		If Not IsBool($bOutlineSymbols) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oCurCont.IsOutlineSymbolsSet = $bOutlineSymbols
		$iError = ($oCurCont.IsOutlineSymbolsSet() = $bOutlineSymbols) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bCharts <> Null) Then
		If Not IsBool($bCharts) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oCurCont.ShowCharts = ($bCharts) ? ($__LOC_ViewObjMode_SHOW) : ($__LOC_ViewObjMode_HIDE)
		$iError = ((($oCurCont.ShowCharts() = $__LOC_ViewObjMode_SHOW) ? (True) : (False)) = $bCharts) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($bDrawing <> Null) Then
		If Not IsBool($bDrawing) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oCurCont.ShowDrawing = ($bDrawing) ? ($__LOC_ViewObjMode_SHOW) : ($__LOC_ViewObjMode_HIDE)
		$iError = ((($oCurCont.ShowDrawing() = $__LOC_ViewObjMode_SHOW) ? (True) : (False)) = $bDrawing) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bObjects <> Null) Then
		If Not IsBool($bObjects) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oCurCont.ShowObjects = ($bObjects) ? ($__LOC_ViewObjMode_SHOW) : ($__LOC_ViewObjMode_HIDE)
		$iError = ((($oCurCont.ShowObjects() = $__LOC_ViewObjMode_SHOW) ? (True) : (False)) = $bObjects) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocViewWindowSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocVisible
; Description ...: Set or retrieve the current visibility of a document.
; Syntax ........: _LOCalc_DocVisible(ByRef $oDoc[, $bVisible = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bVisible            - [optional] Default is Null. If True, the document is visible.
; Return values .: Success: 1 or Boolean.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. $bVisible successfully set.
;                  @Error: 0, @Extended: 1, Return: Boolean = Success. Returning current visibility state of the Document, True if visible, False if invisible.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bVisible not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query whether Document is visible.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOCalc_DocCreate, _LOCalc_DocOpen
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocVisible(ByRef $oDoc, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsVis
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bVisible) Then
		$bIsVis = $oDoc.CurrentController.Frame.ContainerWindow.isVisible()
		If Not IsBool($bIsVis) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $bIsVis)
	EndIf

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.Visible = $bVisible
	$iError = ($oDoc.CurrentController.Frame.ContainerWindow.isVisible() = $bVisible) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocWindowFirstColumn
; Description ...: Set or Retrieve the first visible Column in the Document view.
; Syntax ........: _LOCalc_DocWindowFirstColumn(ByRef $oDoc[, $iColumn = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iColumn             - [optional] Default is Null. The column number to set as the first visible column on the page, 0 based.
; Return values .: Success: 1 or Integer
;                  @Error: 0, @Extended: 0, Return: 1 = Success. First visible Column was successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning the first visible column number as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iColumn not an Integer, less than 0 or greater than number of columns contained in the document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve the First Column value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iColumn
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  This will fail if there are currently any frozen Columns.
; Related .......: _LOCalc_DocWindowFirstRow, _LOCalc_DocColumnsRowsFreeze, _LOCalc_DocWindowVisibleRange, _LOCalc_DocWindowSplit
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocWindowFirstColumn(ByRef $oDoc, $iColumn = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurCol

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iColumn) Then
		$iCurCol = $oDoc.CurrentController.getFirstVisibleColumn()
		If Not IsInt($iCurCol) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurCol)
	EndIf

	If Not __LO_IntIsBetween($iColumn, 0, $oDoc.CurrentController.getActiveSheet().Columns.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.setFirstVisibleColumn($iColumn)
	$iError = ($oDoc.CurrentController.getFirstVisibleColumn() = $iColumn) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocWindowFirstColumn

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocWindowFirstRow
; Description ...: Set or Retrieve the first visible Row in the Document view.
; Syntax ........: _LOCalc_DocWindowFirstRow(ByRef $oDoc[, $iRow = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iRow                - [optional] Default is Null. The row number to set as the first visible row on the page, 0 based.
; Return values .: Success: 1 or Integer
;                  @Error: 0, @Extended: 0, Return: 1 = Success. First visible Row was successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning the first visible row number as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iRow not an Integer, less than 0 or greater than number of Rows contained in the document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve the First Row value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iRow
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  This will fail if there are currently any frozen Rows.
; Related .......: _LOCalc_DocWindowFirstColumn, _LOCalc_DocColumnsRowsFreeze, _LOCalc_DocWindowVisibleRange, _LOCalc_DocWindowSplit
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocWindowFirstRow(ByRef $oDoc, $iRow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurRow

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iRow) Then
		$iCurRow = $oDoc.CurrentController.getFirstVisibleRow()
		If Not IsInt($iCurRow) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurRow)
	EndIf

	If Not __LO_IntIsBetween($iRow, 0, $oDoc.CurrentController.getActiveSheet().Rows.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.setFirstVisibleRow($iRow)
	$iError = ($oDoc.CurrentController.getFirstVisibleRow() = $iRow) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocWindowFirstRow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocWindowIsSplit
; Description ...: Query whether the current Document's view is split.
; Syntax ........: _LOCalc_DocWindowIsSplit(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Document view is currently split.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query Document whether the Document view is currently split.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocWindowSplit, _LOCalc_DocColumnsRowsAreFrozen
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocWindowIsSplit(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bReturn = $oDoc.CurrentController.getIsWindowSplit()
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOCalc_DocWindowIsSplit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocWindowSplit
; Description ...: Split a Document's View either Horizontally, Vertically, or both, or retrieve the current split settings.
; Syntax ........: _LOCalc_DocWindowSplit(ByRef $oDoc[, $iX = Null[, $iY = Null[, $bReturnPixels = True]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iX                  - [optional] Default is Null. See remarks. The Horizontal (X) position to split the View, in pixels. Call with 0 for no Horizontal split.
;                  $iY                  - [optional] Default is Null. See remarks. The Vertical (Y) position to split the View, in pixels. Call with 0 to skip.
;                  $bReturnPixels       - [optional] Default is True. See remarks. If True, return value will be in pixels, Else, return value will be Column Number (X), and Row Number (Y).
; Return values .: Success: 1, 2 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in Pixels, in a 2 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: 2, Return: Array = Success. All optional parameters were called with Null, returning current settings in Column/Row values, in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bReturnPixels not a Boolean.
;                  @Error: 1, @Extended: 3 = $iX not an Integer.
;                  @Error: 1, @Extended: 4 = $iY not an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To remove the split view, set both $iX and $iY to 0.
;                  $bReturnPixels changes only the return value type, it doesn't change the type of input values to use for $iX and $iY.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOCalc_DocWindowIsSplit, _LOCalc_DocWindowFirstColumn, _LOCalc_DocWindowFirstRow, _LOCalc_DocColumnsRowsFreeze
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocWindowSplit(ByRef $oDoc, $iX = Null, $iY = Null, $bReturnPixels = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aiWindow[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnPixels) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iX, $iY) Then
		If $bReturnPixels Then
			__LO_ArrayFill($aiWindow, $oDoc.CurrentController.getSplitHorizontal(), $oDoc.CurrentController.getSplitVertical())

			Return SetError($__LO_STATUS_SUCCESS, 1, $aiWindow)

		Else
			__LO_ArrayFill($aiWindow, $oDoc.CurrentController.getSplitColumn(), $oDoc.CurrentController.getSplitRow())

			Return SetError($__LO_STATUS_SUCCESS, 2, $aiWindow)
		EndIf
	EndIf

	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDoc.CurrentController.splitAtPosition($iX, $iY)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocWindowSplit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocWindowVisibleRange
; Description ...: Retrieve a Cell Range Object for the currently visible cells in the document view.
; Syntax ........: _LOCalc_DocWindowVisibleRange(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning currently visible Range of cells as a Cell Range Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve currently visible Range Address.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Sheet Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Cell Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocWindowFirstColumn, _LOCalc_DocWindowFirstRow, _LOCalc_DocWindowSplit, _LOCalc_DocColumnsRowsFreeze
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocWindowVisibleRange(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tRange
	Local $oSheet, $oRange

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tRange = $oDoc.CurrentController.getVisibleRange()
	If Not IsObj($tRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSheet = $oDoc.Sheets.getByIndex($tRange.Sheet())
	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oRange = $oSheet.getCellRangeByPosition($tRange.StartColumn(), $tRange.StartRow(), $tRange.EndColumn(), $tRange.EndRow())
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oRange)
EndFunc   ;==>_LOCalc_DocWindowVisibleRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocZoom
; Description ...: Modify the zoom value for a document.
; Syntax ........: _LOCalc_DocZoom(ByRef $oDoc[, $iZoomType = Null[, $iZoom = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iZoomType           - [optional] (0-4) Default is Null. The Zoom type, See remarks. See constants $LOC_ZOOMTYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iZoom               - [optional] (20-600) Default is Null. The zoom percentage. Only valid if Zoom type is set to "By Value"
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iZoomType not an Integer, less than 0 or greater than 4. See constants $LOC_ZOOMTYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iZoom not an Integer, less than 20 or greater than 600.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iZoom
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Zoom type always has the value of $LOC_ZOOMTYPE_BY_VALUE(3), when using the other zoom types, the value stays the same, but the zoom level is modified. Consequently, I have not added an error check for the Zoom Type property being correctly set.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOCalc_DocViewWindowSettings, _LOCalc_DocMaximize, _LOCalc_DocMinimize, _LOCalc_DocWindowVisibleRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocZoom(ByRef $oDoc, $iZoomType = Null, $iZoom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiZoom[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iZoomType, $iZoom) Then
		__LO_ArrayFill($aiZoom, $oDoc.CurrentController.ZoomType(), $oDoc.CurrentController.ZoomValue())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiZoom)
	EndIf

	If ($iZoomType <> Null) Then
		If Not __LO_IntIsBetween($iZoomType, $LOC_ZOOMTYPE_OPTIMAL, $LOC_ZOOMTYPE_PAGE_WIDTH_EXACT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.CurrentController.ZoomType = $iZoomType
	EndIf

	If ($iZoom <> Null) Then
		If Not __LO_IntIsBetween($iZoom, 20, 600) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.CurrentController.ZoomValue = $iZoom
		$iError = ($oDoc.CurrentController.ZoomValue() = $iZoom) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocZoom
