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
; _LOCalc_DocPrintersAltGetNames
; _LOCalc_DocPrintersGetNames
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bSaveChanges        - [optional] a boolean value. Default is True. If true, saves changes if any were made before closing. See remarks.
;                  $sSaveName           - [optional] a string value. Default is "". The file name to save the file as, if the file hasn't been saved before. See Remarks.
;                  $bDeliverOwnership   - [optional] a boolean value. Default is True. If True, deliver ownership of the document Object from the script to LibreOffice, recommended is True.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bSaveChanges not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sSaveName not a String.
;                  @Error 1 @Extended 4 Return 0 = $bDeliverOwnership not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Path Conversion to L.O. URL Failed.
;                  @Error 3 @Extended 2 Return 0 = Error while retrieving FilterName.
;                  @Error 3 @Extended 3 Return 0 = Error while setting Filter Name properties.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Success, Document was successfully closed, and was saved to the returned file Path.
;                  @Error 0 @Extended 2 Return String = Success, Document was successfully closed, document's changes were saved to its existing location.
;                  @Error 0 @Extended 3 Return String = Success, Document was successfully closed, document either had no changes to save, or $bSaveChanges was set to False. If document had a save location, or if document was saved to a location, it is returned, else an empty string is returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bSaveChanges is true and the document hasn't been saved yet, the document is saved to the desktop.
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
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	EndIf

	If ($bSaveChanges = True) Then
		If $oDoc.hasLocation() Then
			$oDoc.store()
			$sDocPath = _LO_PathConvert($oDoc.getURL(), $LO_PATHCONV_PCPATH_RETURN)
			$oDoc.Close($bDeliverOwnership)

			Return SetError($__LO_STATUS_SUCCESS, 2, $sDocPath)

		Else
			$oDoc.storeAsURL($sSavePath, $aArgs)
			$oDoc.Close($bDeliverOwnership)

			Return SetError($__LO_STATUS_SUCCESS, 1, _LO_PathConvert($sSavePath, $LO_PATHCONV_PCPATH_RETURN))
		EndIf
	EndIf

	If $oDoc.hasLocation() Then $sDocPath = _LO_PathConvert($oDoc.getURL(), $LO_PATHCONV_PCPATH_RETURN)
	$oDoc.Close($bDeliverOwnership)

	Return SetError($__LO_STATUS_SUCCESS, 3, $sDocPath)
EndFunc   ;==>_LOCalc_DocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocColumnsRowsAreFrozen
; Description ...: Query whether the Document has Columns or Rows currently frozen in view.
; Syntax ........: _LOCalc_DocColumnsRowsAreFrozen(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query Document whether frozen Columns/Rows are present.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the Document currently contains frozen Columns/Rows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iColumns            - [optional] an integer value. Default is 0. The number of Columns to freeze. Set to 0 to skip. See remarks.
;                  $iRows               - [optional] an integer value. Default is 0. The number of Rows to freeze. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumns not an Integer, less than 0 or greater than number of columns contained in the document.
;                  @Error 1 @Extended 3 Return 0 = $iRows not an Integer, less than 0 or greater than number of rows contained in the document.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Called Columns/Rows were successfully frozen.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To set no Columns or rows to be frozen in view, set both $iColumns and $iRows to 0.
;                  Setting either $iColumns or $iRows will lose previous values for both.
; Related .......:
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
; Syntax ........: _LOCalc_DocConnect($sFile[, $bConnectCurrent = False[, $bConnectAll = False]])
; Parameters ....: $sFile               - a string value. A Full or partial file path, or a full or partial file name. See remarks. Can be an empty string if $bConnectAll or $bConnectCurrent is True.
;                  $bConnectCurrent     - [optional] a boolean value. Default is False. If True, returns the currently active, or last active Document, unless it is not a Calc Document.
;                  $bConnectAll         - [optional] a boolean value. Default is False. If True, returns an array containing all open LibreOffice Calc Documents. See remarks.
; Return values .: Success: Object or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFile not a string.
;                  @Error 1 @Extended 2 Return 0 = $bConnectCurrent not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bConnectAll not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating ServiceManager object.
;                  @Error 2 @Extended 2 Return 0 = Error creating Desktop object.
;                  @Error 2 @Extended 3 Return 0 = Error creating enumeration of open documents.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = No open Libre Office documents found.
;                  @Error 3 @Extended 2 Return 0 = Current Component not a Calc Document.
;                  @Error 3 @Extended 3 Return 0 = Error converting path to Libre Office URL.
;                  @Error 3 @Extended 4 Return 0 = No matches found.
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Success, The Object for the current, or last active document is returned.
;                  @Error 0 @Extended ? Return Array = Success, An Array of all open LibreOffice Calc documents is returned. See remarks. @Extended is set to number of results.
;                  @Error 0 @Extended 3 Return Object = Success, The Object for the document with matching URL is returned.
;                  @Error 0 @Extended 4 Return Object = Success, The Object for the document with matching Title is returned.
;                  @Error 0 @Extended 5 Return Object = Success, A partial Title or Path search found only one match, returning the Object for the found document.
;                  @Error 0 @Extended ? Return Array = Success, An Array of all matching Libre Text documents from a partial Title or Path search. See remarks. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFile can be either the full Path (Name and extension included; e.g: C:\file\Test.ods Or file:///C:/file/Test.ods) of the document, or the full Title with extension, (e.g: Test.ods), or a partial file path (e.g: file1\file2\Test Or file1\file2 Or file1/file2/ etc.), or a partial name (e.g: test, would match Test1.ods, Test2.xlsx etc.).
;                  Partial file path searches and file name searches, as well as the connect all option, return arrays with three columns per result. ($aArray[0][3]). each result is stored in a separate row;
;                  Row 1, Column 0 contains the Object for that document. e.g. $aArray[0][0] = $oDoc
;                  Row 1, Column 1 contains the Document's full title and extension. e.g. $aArray[0][1] = This Test File.xlsx
;                  Row 1, Column 2 contains the document's full file path. e.g. $aArray[0][2] = C:\Folder1\Folder2\This Test File.xlsx
;                  Row 2, Column 0 contains the Object for the next document. And so on. e.g. $aArray[1][0] = $oDoc2
; Related .......: _LOCalc_DocOpen, _LOCalc_DocClose, _LOCalc_DocCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocConnect($sFile, $bConnectCurrent = False, $bConnectAll = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local Const $__STR_STRIPLEADING = 1
	Local $aoConnectAll[1], $aoPartNameSearch[1]
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop
	Local $sServiceName = "com.sun.star.sheet.SpreadsheetDocument"

	If Not IsString($sFile) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bConnectCurrent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bConnectAll) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	If Not $oDesktop.getComponents.hasElements() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; no L.O open

	$oEnumDoc = $oDesktop.getComponents.createEnumeration()
	If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If $bConnectCurrent Then
		$oDoc = $oDesktop.currentComponent()

		Return ($oDoc.supportsService($sServiceName)) ? (SetError($__LO_STATUS_SUCCESS, 1, $oDoc)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0))
	EndIf

	If $bConnectAll Then
		ReDim $aoConnectAll[1][3]
		$iCount = 0
		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService($sServiceName) Then
				ReDim $aoConnectAll[$iCount + 1][3]
				$aoConnectAll[$iCount][0] = $oDoc
				$aoConnectAll[$iCount][1] = $oDoc.Title()
				$aoConnectAll[$iCount][2] = _LO_PathConvert($oDoc.getURL(), $LO_PATHCONV_PCPATH_RETURN)
				$iCount += 1
			EndIf
			Sleep(10)
		WEnd

		Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoConnectAll)
	EndIf

	$sFile = StringStripWS($sFile, $__STR_STRIPLEADING)
	If StringInStr($sFile, "\") Then $sFile = _LO_PathConvert($sFile, $LO_PATHCONV_OFFICE_RETURN) ; Convert to L.O File path.
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If StringInStr($sFile, "file:///") Then ; URL/Path and Name search

		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()

			If ($oDoc.getURL() == $sFile) Then Return SetError($__LO_STATUS_SUCCESS, 3, $oDoc) ; Match
		WEnd

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; no match

	Else
		If Not StringInStr($sFile, "/") And StringInStr($sFile, ".") Then ; Name with extension only search
			While $oEnumDoc.hasMoreElements()
				$oDoc = $oEnumDoc.nextElement()
				If StringInStr($oDoc.Title, $sFile) Then Return SetError($__LO_STATUS_SUCCESS, 4, $oDoc) ; Match
			WEnd

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; no match
		EndIf

		$iCount = 0 ; partial name or partial url search
		ReDim $aoPartNameSearch[$iCount + 1][3]

		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If StringInStr($sFile, "/") Then
				If StringInStr($oDoc.getURL(), $sFile) Then
					ReDim $aoPartNameSearch[$iCount + 1][3]
					$aoPartNameSearch[$iCount][0] = $oDoc
					$aoPartNameSearch[$iCount][1] = $oDoc.Title
					$aoPartNameSearch[$iCount][2] = _LO_PathConvert($oDoc.getURL, $LO_PATHCONV_PCPATH_RETURN)
					$iCount += 1
				EndIf

			Else
				If StringInStr($oDoc.Title, $sFile) Then
					ReDim $aoPartNameSearch[$iCount + 1][3]
					$aoPartNameSearch[$iCount][0] = $oDoc
					$aoPartNameSearch[$iCount][1] = $oDoc.Title
					$aoPartNameSearch[$iCount][2] = _LO_PathConvert($oDoc.getURL, $LO_PATHCONV_PCPATH_RETURN)
					$iCount += 1
				EndIf
			EndIf
		WEnd
		If IsString($aoPartNameSearch[0][1]) Then
			If (UBound($aoPartNameSearch) = 1) Then

				Return SetError($__LO_STATUS_SUCCESS, 5, $aoPartNameSearch[0][0]) ; matches

			Else

				Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoPartNameSearch) ; matches
			EndIf

		Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; no match
		EndIf
	EndIf
EndFunc   ;==>_LOCalc_DocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocCreate
; Description ...: Open a new Libre Office Calc Document or Connect to an existing blank, unsaved, writable document.
; Syntax ........: _LOCalc_DocCreate([$bForceNew = True[, $bHidden = False]])
; Parameters ....: $bForceNew           - [optional] a boolean value. Default is True. If True, force opening a new Calc Document instead of checking for a usable blank.
;                  $bHidden             - [optional] a boolean value. Default is False. If True opens the new document invisible or changes the existing document to invisible.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bForceNew not a Boolean.
;                  @Error 1 @Extended 2 Return 0 = $bHidden not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failure Creating Object com.sun.star.ServiceManager.
;                  @Error 2 @Extended 2 Return 0 = Failure Creating Object com.sun.star.frame.Desktop.
;                  @Error 2 @Extended 3 Return 0 = Failed to enumerate available documents.
;                  @Error 2 @Extended 4 Return 0 = Failure Creating New Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Document Object is still returned. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHidden
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Successfully connected to an existing Document. Returning Document's Object
;                  @Error 0 @Extended 2 Return Object = Successfully created a new document. Returning Document's Object
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sFilePath           - a string value. Full path to save the document to, including Filename and extension. See Remarks.
;                  $bSamePath           - [optional] a boolean value. Default is False. If True, uses the path of the current document to export to. See Remarks
;                  $sFilterName         - [optional] a string value. Default is "". Filter name. If set to "" (blank string), Filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .ods extension is used instead, with the filter name of "calc8".
;                  $bOverwrite          - [optional] a boolean value. Default is Null. If True, file will be overwritten.
;                  $sPassword           - [optional] a string value. Default is Null. Password String to set for the document. (Not all file formats can have a Password set). "" (blank string) or Null = No Password.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFilePath not a String.
;                  @Error 1 @Extended 3 Return 0 = $bSamePath not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $sFilterName not a String.
;                  @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $sPassword not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Converting Path to/from L.O. URL
;                  @Error 3 @Extended 2 Return 0 = Document has no save path, and $bSamePath is set to True.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving FilterName.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Error setting FilterName Property
;                  @Error 4 @Extended 2 Return 0 = Error setting Overwrite Property
;                  @Error 4 @Extended 3 Return 0 = Error setting Password Property
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning save path for exported document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Does not alter the original save path (if there was one), saves a copy of the document to the new path, in the new file format if one is chosen.
;                  If $bSamePath is set to True, the same save path as the current document is used. You must still fill in "$sFilePath" with the desired File Name and new extension, but you do not need to enter the file path.
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
	If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0)
	EndIf

	If ($sPassword <> Null) Then
		If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0)
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iHeight             - [optional] an integer value (1-25). Default is Null. The number of lines to display in the formula bar.
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iHeight not an Integer, less than 1 or greater than 25.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iHeight
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version less than 7.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current Formula Bar Height as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocFormulaBarHeight(ByRef $oDoc, $iHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.4) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	If __LO_VarsAreNull($iHeight) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.FormulaBarHeight())

	If Not __LO_IntIsBetween($iHeight, 1, 25) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.FormulaBarHeight = $iHeight
	$iError = ($oDoc.CurrentController.FormulaBarHeight() = $iHeight) ? ($iError) : (BitOR($iError, 1))

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc   ;==>_LOCalc_DocFormulaBarHeight

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocGetName
; Description ...: Retrieve the document's name.
; Syntax ........: _LOCalc_DocGetName(ByRef $oDoc[, $bReturnFull = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bReturnFull         - [optional] a boolean value. Default is False. If True, the full window title is returned, such as is used by AutoIt window related functions.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bReturnFull not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returns the document's current Name/Title
;                  @Error 0 @Extended 1 Return String = Success. Returns the document's current Window Title, which includes the document name and usually: "-LibreOffice Calc".
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
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

	$sName = ($bReturnFull = True) ? ($oDoc.CurrentController.Frame.Title()) : ($oDoc.Title())

	Return ($bReturnFull = True) ? (SetError($__LO_STATUS_SUCCESS, 1, $sName)) : (SetError($__LO_STATUS_SUCCESS, 0, $sName))
EndFunc   ;==>_LOCalc_DocGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocGetPath
; Description ...: Returns a Document's current save path.
; Syntax ........: _LOCalc_DocGetPath(ByRef $oDoc[, $bReturnLibreURL = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bReturnLibreURL     - [optional] a boolean value. Default is False. If True, returns a path in Libre Office URL format, else false returns a regular Windows path.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bReturnLibreURL not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = Document has no save path.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error converting Libre URL to Computer path format.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returns the P.C. path to the current document's save path.
;                  @Error 0 @Extended 1 Return String = Success. Returns the Libre Office URL to the current document's save path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LO_PathConvert
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

	Return ($bReturnLibreURL = True) ? (SetError($__LO_STATUS_SUCCESS, 1, $sPath)) : (SetError($__LO_STATUS_SUCCESS, 0, $sPath))
EndFunc   ;==>_LOCalc_DocGetPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocHasPath
; Description ...: Returns whether a document has been saved to a location already or not.
; Syntax ........: _LOCalc_DocHasPath(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the document has a save location. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocHasPath(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.hasLocation())
EndFunc   ;==>_LOCalc_DocHasPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocIsActive
; Description ...: Tests if called document is the active document of other Libre windows.
; Syntax ........: _LOCalc_DocIsActive(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if document is the currently active Libre window. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This does NOT test if the document is the current active window in Windows, it only tests if the document is the current active document among other Libre Office documents.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocIsActive(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.CurrentController.Frame.isActive())
EndFunc   ;==>_LOCalc_DocIsActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocIsModified
; Description ...: Test whether the document has been modified since being created or since the last save.
; Syntax ........: _LOCalc_DocIsModified(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the document has been modified since last being saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocIsModified(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.isModified())
EndFunc   ;==>_LOCalc_DocIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocIsReadOnly
; Description ...: Tests whether a document is currently set to Read Only.
; Syntax ........: _LOCalc_DocIsReadOnly(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True is document is currently Read Only, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only documents that have been saved to a location, will ever be "ReadOnly".
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocIsReadOnly(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.isReadOnly())
EndFunc   ;==>_LOCalc_DocIsReadOnly

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocMaximize
; Description ...: Maximize or restore a document.
; Syntax ........: _LOCalc_DocMaximize(ByRef $oDoc[, $bMaximize = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bMaximize           - [optional] a boolean value. Default is Null. If True, document window is maximized, else if false, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bMaximize not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully maximized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMaximize set to Null, returning boolean indicating if Document is currently maximized (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bMaximize is set to Null, returns a Boolean indicating if document is currently maximized (True).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocMaximize(ByRef $oDoc, $bMaximize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bMaximize = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.IsMaximized())

	If Not IsBool($bMaximize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.IsMaximized = $bMaximize

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocMaximize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocMinimize
; Description ...: Minimize or restore a document.
; Syntax ........: _LOCalc_DocMinimize(ByRef $oDoc[, $bMinimize = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bMinimize           - [optional] a boolean value. Default is Null. If True, document window is minimized, else if false, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bMinimize not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully minimized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMinimize set to Null, returning boolean indicating if Document is currently minimized (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bMinimize is set to Null, returns a Boolean indicating if document is currently minimized (True).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocMinimize(ByRef $oDoc, $bMinimize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bMinimize = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.IsMinimized())

	If Not IsBool($bMinimize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.IsMinimized = $bMinimize

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocMinimize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocOpen
; Description ...: Open an existing Calc Document, returning its object identifier.
; Syntax ........: _LOCalc_DocOpen($sFilePath[, $bConnectIfOpen = True[, $bHidden = Null[, $bReadOnly = Null[, $sPassword = Null[, $bLoadAsTemplate = Null[, $sFilterName = Null]]]]]])
; Parameters ....: $sFilePath           - a string value. Full path and filename of the file to be opened.
;                  $bConnectIfOpen      - [optional] a boolean value. Default is True(Connect). Whether to connect to the requested document if it is already open. See remarks.
;                  $bHidden             - [optional] a boolean value. Default is Null. If true, opens the document invisibly.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If true, opens the document as read-only.
;                  $sPassword           - [optional] a string value. Default is Null. The password that was used to read-protect the document, if any.
;                  $bLoadAsTemplate     - [optional] a boolean value. Default is Null. If true, opens the document as a Template, i.e. an untitled copy of the specified document is made instead of modifying the original document.
;                  $sFilterName         - [optional] a string value. Default is Null. Name of a LibreOffice filter to use to load the specified document. LibreOffice automatically selects which to use by default.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFilePath not string, or file not found.
;                  @Error 1 @Extended 2 Return 0 = Error converting file path to URL path.
;                  @Error 1 @Extended 3 Return 0 = $bConnectIfOpen not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $sPassword not a string.
;                  @Error 1 @Extended 7 Return 0 = $bLoadAsTemplate not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $sFilterName not a string.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create ServiceManager Object
;                  @Error 2 @Extended 2 Return 0 = Failed to create Desktop Object
;                  @Error 2 @Extended 3 Return 0 = Failed opening or connecting to document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHidden
;                  |                               2 = Error setting $bReadOnly
;                  |                               4 = Error setting $sPassword
;                  |                               8 = Error setting $bLoadAsTemplate
;                  |                               16 = Error setting $sFilterName
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Successfully connected to requested Document without requested parameters. Returning Document's Object.
;                  @Error 0 @Extended 2 Return Object = Successfully opened requested Document with requested parameters. Returning Document's Object.
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

	If $bConnectIfOpen Then $oDoc = _LOCalc_DocConnect($sFilePath)
	If IsObj($oDoc) Then Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))

	$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOCalc_DocOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocPosAndSize
; Description ...: Reposition and resize a document window.
; Syntax ........: _LOCalc_DocPosAndSize(ByRef $oDoc[, $iX = Null[, $iY = Null[, $iWidth = Null[, $iHeight = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iX                  - [optional] an integer value. Default is Null. The X coordinate of the window.
;                  $iY                  - [optional] an integer value. Default is Null. The Y coordinate of the window.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the window, in pixels(?).
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the window, in pixels(?).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iY not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iHeight not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Position and Size Structure Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Position and Size Structure Object for error checking.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iX
;                  |                               2 = Error setting $iY
;                  |                               4 = Error setting $iWidth
;                  |                               8 = Error setting $iHeight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: X & Y, on my computer at least, seem to go no lower than 8(X) and 30(Y), if you enter lower than this, it will cause a "property setting Error".
;                  If you want more accurate functionality, use the "WinMove" AutoIt function.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
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

	$iError = ($iX = Null) ? ($iError) : (($tWindowSize.X() = $iX) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iY = Null) ? ($iError) : (($tWindowSize.Y() = $iY) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iWidth = Null) ? ($iError) : (($tWindowSize.Width() = $iWidth) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iHeight = Null) ? ($iError) : (($tWindowSize.Height() = $iHeight) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc   ;==>_LOCalc_DocPosAndSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocPrint
; Description ...: Print a document using the specified settings.
; Syntax ........: _LOCalc_DocPrint(ByRef $oDoc[, $iCopies = 1[, $bCollate = True[, $vPages = "ALL"[, $bWait = True[, $iDuplexMode = $LOC_DUPLEX_OFF[, $sPrinter = ""[, $sFilePathName = ""]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iCopies             - [optional] an integer value. Default is 1. Specifies the number of copies to print.
;                  $bCollate            - [optional] a boolean value. Default is True. Advises the printer to collate the pages of the copies.
;                  $vPages              - [optional] a String or Integer value. Default is "ALL". Specifies which pages to print. See remarks.
;                  $bWait               - [optional] a boolean value. Default is True. If True, the corresponding print request will be executed synchronous. Default is to use synchronous print mode.
;                  $iDuplexMode         - [optional] an integer value (0-3). Default is $__g_iDuplexOFF. Determines the duplex mode for the print job. See Constants, $LOC_DUPLEX_* as defined in LibreOfficeCalc_Constants.au3.
;                  $sPrinter            - [optional] a string value. Default is "". Printer name. If left blank, or if printer name is not found, default printer is used.
;                  $sFilePathName       - [optional] a string value. Default is "". Specifies the name of a file to print to. Creates a .prn file at the given Path. Must include the desired path destination with file name.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iCopies not a Integer.
;                  @Error 1 @Extended 3 Return 0 = $bCollate not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $vPages not an Integer or String.
;                  @Error 1 @Extended 5 Return 0 = $vPages contains invalid characters, a-z, or a period(.).
;                  @Error 1 @Extended 6 Return 0 = $bWait not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iDuplexMode not an Integer, less than 0 or greater than 3. See Constants, $LOC_DUPLEX_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $sPrinter not a String.
;                  @Error 1 @Extended 9 Return 0 = $sFilePathName not a String.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Error setting "Printer Name".
;                  @Error 4 @Extended 2 Return 0 = Error setting "Copies".
;                  @Error 4 @Extended 3 Return 0 = Error setting "Collate".
;                  @Error 4 @Extended 4 Return 0 = Error setting "Wait".
;                  @Error 4 @Extended 5 Return 0 = Error setting "DuplexMode".
;                  @Error 4 @Extended 6 Return 0 = Error setting "Pages".
;                  @Error 4 @Extended 7 Return 0 = Error converting PrintToFile Path.
;                  @Error 4 @Extended 8 Return 0 = Error setting "PrintToFile".
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success Document was successfully printed.
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
Func _LOCalc_DocPrint(ByRef $oDoc, $iCopies = 1, $bCollate = True, $vPages = "ALL", $bWait = True, $iDuplexMode = $LOC_DUPLEX_OFF, $sPrinter = "", $sFilePathName = "")
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
	If Not __LO_IntIsBetween($iDuplexMode, $LOC_DUPLEX_OFF, $LOC_DUPLEX_SHORT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not IsString($sPrinter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	$sPrinter = StringStripWS(StringStripWS($sPrinter, $__STR_STRIPTRAILING), $__STR_STRIPLEADING)
	If Not IsString($sFilePathName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	$sFilePathName = StringStripWS(StringStripWS($sFilePathName, $__STR_STRIPTRAILING), $__STR_STRIPLEADING)
	If $sPrinter <> "" Then
		$asSetPrinterOpt[0] = __LO_SetPropertyValue("Name", $sPrinter)
		If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

		$oDoc.setPrinter($asSetPrinterOpt)
	EndIf
	$avPrintOpt[0] = __LO_SetPropertyValue("CopyCount", $iCopies)
	If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0)

	$avPrintOpt[1] = __LO_SetPropertyValue("Collate", $bCollate)
	If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0)

	$avPrintOpt[2] = __LO_SetPropertyValue("Wait", $bWait)
	If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 4, 0)

	$avPrintOpt[3] = __LO_SetPropertyValue("DuplexMode", $iDuplexMode)
	If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 5, 0)

	If $vPages <> "ALL" Then
		ReDim $avPrintOpt[UBound($avPrintOpt) + 1]
		$avPrintOpt[UBound($avPrintOpt) - 1] = __LO_SetPropertyValue("Pages", $vPages)
		If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 6, 0)
	EndIf
	If $sFilePathName <> "" Then
		$sFilePathName = $sFilePathName & ".prn"
		$sFilePathName = _LO_PathConvert($sFilePathName, $LO_PATHCONV_OFFICE_RETURN)
		If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 7, 0)

		ReDim $avPrintOpt[UBound($avPrintOpt) + 1]
		$avPrintOpt[UBound($avPrintOpt) - 1] = __LO_SetPropertyValue("FileName", $sFilePathName)
		If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 8, 0)
	EndIf
	$oDoc.print($avPrintOpt)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_DocPrint

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocPrintersAltGetNames
; Description ...: Alternate function; Retrieve an array of names for all installed printers, or the current default printer.
; Syntax ........: _LOCalc_DocPrintersAltGetNames([$sPrinterName = ""[, $bReturnDefault = False]])
; Parameters ....: $sPrinterName        - [optional] a string value. Default is "". Name of the printer to list. Default "" returns the list of all printers. See remarks.
;                  $bReturnDefault      - [optional] a boolean value. Default is False. If True, returns only the name of the current default printer.
; Return values .: Success: Array or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sPrinterName not a String.
;                  @Error 1 @Extended 2 Return 0 = $bReturnDefault not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failure Creating Object.
;                  @Error 2 @Extended 2 Return 0 = Failure retrieving printer list Object.
;                  --Printer Related Errors--
;                  @Error 5 @Extended 1 Return 0 = No default printer found.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Returning the default printer name. See remarks.
;                  @Error 0 @Extended ? Return Array = Returning an array of strings containing all installed printers. See remarks. @Extended is set to the number of results.
; Author ........: jguinch (_PrintMgr_EnumPrinter)
; Modified ......: donnyh13 - Added input error checking. Added a return default printer only option.
; Remarks .......: When $bReturnDefault is False, The function returns all installed printers for the user running the script in an array.
;                  If $sPrinterName is set, the name must be exact or no results will be found, unless you use an asterisk (*) for partial name searches, either prefixed (*Canon), suffixed (Canon*), or both (*Canon*).
;                  When $bReturnDefault is True, The function returns only the default printer's name or sets an error if no default printer is found.
; Related .......: _LOCalc_DocPrintersGetNames
; Link ..........: https://www.autoitscript.com/forum/topic/155485-printers-management-udf/
; UDF title......: Printmgr.au3
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocPrintersAltGetNames($sPrinterName = "", $bReturnDefault = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asPrinterNames[10]
	Local $sFilter
	Local $iCount = 0
	Local Const $wbemFlagReturnImmediately = 0x10, $wbemFlagForwardOnly = 0x20
	Local $oWMIService, $oPrinters

	If Not IsString($sPrinterName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $sPrinterName <> "" Then $sFilter = StringReplace(" Where Name like '" & StringReplace($sPrinterName, "\", "\\") & "'", "*", "%")
	$oWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	If Not IsObj($oWMIService) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oPrinters = $oWMIService.ExecQuery("Select * from Win32_Printer" & $sFilter, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If Not IsObj($oPrinters) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	For $oPrinter In $oPrinters
		Switch $bReturnDefault
			Case False
				If $iCount >= (UBound($asPrinterNames) - 1) Then ReDim $asPrinterNames[UBound($asPrinterNames) * 2]
				$asPrinterNames[$iCount] = $oPrinter.Name
				$iCount += 1

			Case True
				If $oPrinter.Default Then Return SetError($__LO_STATUS_SUCCESS, 1, $oPrinter.Name)
		EndSwitch
	Next
	If $bReturnDefault Then Return SetError($__LO_STATUS_PRINTER_RELATED_ERROR, 1, 0)

	ReDim $asPrinterNames[$iCount]

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $asPrinterNames)
EndFunc   ;==>_LOCalc_DocPrintersAltGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocPrintersGetNames
; Description ...: Retrieve an array of names for all installed printers, or the current default printer.
; Syntax ........: _LOCalc_DocPrintersGetNames([$bDefaultOnly = False])
; Parameters ....: $bDefaultOnly        - [optional] a boolean value. Default is False. If True, returns only the name of the current default printer. Libre 6.3 and up only.
; Return values .: Success: An array or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bDefaultOnly Not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failure Creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failure creating "com.sun.star.awt.PrinterServer" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Default printer name.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Array of printer names.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.1.
;                  @Error 6 @Extended 2 Return 0 = Current Libre Office version lower than 6.3.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Returning the default printer name.
;                  @Error 0 @Extended ? Return Array = Returning an array of strings containing all installed printers. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function works for LibreOffice 4.1 and Up.
; Related .......: _LOCalc_DocPrintersAltGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocPrintersGetNames($bDefaultOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oPrintServer
	Local $sDefault
	Local $asPrinters[0]

	If Not __LO_VersionCheck(4.1) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsBool($bDefaultOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = __LO_ServiceManager()
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oPrintServer = $oServiceManager.createInstance("com.sun.star.awt.PrinterServer")
	If Not IsObj($oPrintServer) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If $bDefaultOnly Then
		If Not __LO_VersionCheck(6.3) Then Return SetError($__LO_STATUS_VER_ERROR, 2, 0)

		$sDefault = $oPrintServer.getDefaultPrinterName()
		If IsString($sDefault) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sDefault)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	$asPrinters = $oPrintServer.getPrinterNames()
	If IsArray($asPrinters) Then Return SetError($__LO_STATUS_SUCCESS, UBound($asPrinters), $asPrinters)

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
EndFunc   ;==>_LOCalc_DocPrintersGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocRedo
; Description ...: Perform one Redo action for a document.
; Syntax ........: _LOCalc_DocRedo(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Document does not have a redo action to perform.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Successfully performed a redo action.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocUndo, _LOCalc_DocRedoIsPossible, _LOCalc_DocRedoGetAllActionTitles, _LOCalc_DocRedoCurActionTitle
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared all Redo Actions.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This will silently fail if there are any _LOCalc_DocUndoActionBegin still active.
; Related .......: _LOCalc_DocUndoClear, _LOCalc_DocUndoReset
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Redo Action.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Returning the current available redo action title as a String. Will be an empty String if no action is available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocRedo, _LOCalc_DocRedoGetAllActionTitles
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve an array of Redo action titles.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Returning all available redo action Titles in an array of Strings. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocRedo, _LOCalc_DocRedoCurActionTitle
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the document has a redo action to perform, True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocRedo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocRedoIsPossible(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.UndoManager.isRedoPossible())
EndFunc   ;==>_LOCalc_DocRedoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocSave
; Description ...: Save any changes made to a Document.
; Syntax ........: _LOCalc_DocSave(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Document is Read Only or Document has no save location, try SaveAs.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Document Successfully saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocExport, _LOCalc_DocSaveAs
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sFilePath           - a string value. Full path to save the document to, including Filename and extension.
;                  $sFilterName         - [optional] a string value. Default is "". The filter name. Calling "" (blank string), means the filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .ods extension is used instead, with the filter name of "calc8".
;                  $bOverwrite          - [optional] a boolean value. Default is Null. If True, the existing file will be overwritten.
;                  $sPassword           - [optional] a string value. Default is Null. Sets a password for the document. (Not all file formats can have a Password set). Null or "" (blank string) = No Password.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sFilePath not a String.
;                  @Error 1 @Extended 3 Return 0 = $sFilterName not a String.
;                  @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $sPassword not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Converting Path to/from L.O. URL
;                  @Error 3 @Extended 2 Return 0 = Error retrieving FilterName.
;                  @Error 3 @Extended 3 Return 0 = Error setting FilterName Property
;                  @Error 3 @Extended 4 Return 0 = Error setting Overwrite Property
;                  @Error 3 @Extended 5 Return 0 = Error setting Password Property
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Successfully Saved the document. Returning document save path.
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
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	If $sPassword <> Null Then
		If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	$oDoc.storeAsURL($sFilePath, $aProperties)

	$sSavePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOCalc_DocSaveAs

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocSelectionCopy
; Description ...: "Copies" data selected by the ViewCursor, returning an Object for use in inserting later.
; Syntax ........: _LOCalc_DocSelectionCopy(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Copy Selected Data, make sure Data is selected.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Data was successfully copied, returning an Object for use in _LOCalc_DocSelectionPaste.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Object or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current selection.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve count of multiple selections.
;                  @Error 3 @Extended 3 Return 0 = Failed to determine selection type.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Single cell selected or cursor is editing a cell, returning Cell Object.
;                  @Error 0 @Extended 1 Return Object = Success. Cell Range selected, returning Cell Range Object.
;                  @Error 0 @Extended ? Return Array = Success. Multiple Cells or Cell Ranges selected, returning array of Cell Range Objects.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the user has nothing selected, or is typing in a cell, the return will still be the single cell Object.
; Related .......:
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oData               - [in/out] an object. A Object returned from _LOCalc_DocSelectionCopy to insert.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oData not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Data was successfully inserted at the currently selected cell.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oObj                - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 3 Return 0 = Object called in $oObj not a Cell Object and not a Cell Range.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Object called in $oObj successfully selected.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $aoRange             - [in/out] an array of objects. An array of Cell or Cell Range objects returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an object.
;                  @Error 1 @Extended 2 Return 0 = $aoRange not an Array.
;                  @Error 1 @Extended 3 Return ? = Array called in $aoRange does not contain an Object in returned element number.
;                  @Error 1 @Extended 4 Return ? = Array called in $aoRange contains an Object in returned element that is not a Cell Object and not a Cell Range.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.sheet.SheetCellRanges" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return ? = Failed to retrieve Range Address from Object located in array called in $aoRange, returning problem element.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Objects were successfully selected.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Window was successfully brought to the front of the open windows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If minimized, the document is restored and brought to the front of the visible pages. Generally only brings the document to the front of other Libre Office windows.
; Related .......:
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Document does not have an undo action to perform.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Successfully performed an undo action.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sName               - [optional] a string value. Default is "AU3LO-Automation". The name of the Undo Action to display in the UI when completed.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully began an Undo Action group recording.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This begins an Undo Action Group, any functions and actions done after this function is called will be grouped together, and if undone, all actions will be undone together at once.
;                  _LOCalc_DocUndoActionEnd must be called after this function before this undo group will become available in the Undo Action list.
;                  _LOCalc_DocUndoActionBegin can be nested, e.g. call this function multiple times without ending the first undo action, but only the last group that is ended with _LOCalc_DocUndoActionEnd will appear.
; Related .......:
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully ended the last Undo Action group recording.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This stops the grouping of actions into the last created Undo Action Group.
; Related .......: _LOCalc_DocUndoActionBegin
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared all Undo and Redo Actions.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This will silently fail if there are any _LOCalc_DocUndoActionBegin still active.
; Related .......: _LOCalc_DocRedoClear, _LOCalc_DocUndoReset
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Undo Action.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Returning the current available Undo action title as a String. Will be an empty String if no action is available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocUndo, _LOCalc_DocUndoGetAllActionTitles
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve an array of Undo action titles.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Returning all available undo action Titles in an array of Strings. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocUndo, _LOCalc_DocUndoCurActionTitle
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the document has an undo action to perform, True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_DocUndo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndoIsPossible(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.UndoManager.isUndoPossible())
EndFunc   ;==>_LOCalc_DocUndoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndoReset
; Description ...: Reset the UndoManager.
; Syntax ........: _LOCalc_DocUndoReset(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully reset the undo manager.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling this function does the following: remove all locks from the undo manager; closes all open undo group actions, clears all undo actions, clears all redo actions.
; Related .......: _LOCalc_DocRedoClear, _LOCalc_DocUndoClear
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bFormulas           - [optional] a boolean value. Default is Null. If True, Formulas, rather than results, are displayed in the cells.
;                  $bZeroValues         - [optional] a boolean value. Default is Null. If True, numbers with the value of 0 are shown.
;                  $bComments           - [optional] a boolean value. Default is Null. If True, a small rectangle in the top right corner of the cell indicates that a comment exists.
;                  $bPageBreaks         - [optional] a boolean value. Default is Null. If True, Page Breaks are displayed for a print area.
;                  $bHelpLines          - [optional] a boolean value. Default is Null. If True, help lines are displayed while moving graphics, drawings, etc.
;                  $bValueHighlight     - [optional] a boolean value. Default is Null. If True, Cell contents are displayed in different colors, depending on the content type of the cell.
;                  $bAnchors            - [optional] a boolean value. Default is Null. If True, the Anchor icon is displayed when a graphic or other object is selected.
;                  $bGrid               - [optional] a boolean value. Default is Null. If True, Gridlines are displayed.
;                  $iGridColor          - [optional] an integer value (0-16777215). Default is Null. Set the Grid line color in Long Color format. Can be one of the constants $LO_COLOR_* as defined in LibreOffice_Constants.au3 or a custom value.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bFormulas not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bZeroValues not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bComments not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bPageBreaks not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bHelpLines not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bValueHighlight not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bAnchors not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bGrid not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $iGridColor not an Integer, less than 0 or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Current Controller Object for the Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFormulas
;                  |                               2 = Error setting $bZeroValues
;                  |                               4 = Error setting $bComments
;                  |                               8 = Error setting $bPageBreaks
;                  |                               16 = Error setting $bHelpLines
;                  |                               32 = Error setting $bValueHighlight
;                  |                               64 = Error setting $bAnchors
;                  |                               128 = Error setting $bGrid
;                  |                               256 = Error setting $iGridColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 9 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bHeaders            - [optional] a boolean value. Default is Null. If True, Column/Row headers are displayed.
;                  $bVertScroll         - [optional] a boolean value. Default is Null. If True, a Vertical scrollbar is displayed at the right of the document.
;                  $bHoriScroll         - [optional] a boolean value. Default is Null. If True, a Horizontal scrollbar is displayed at the bottom of the document.
;                  $bSheetTabs          - [optional] a boolean value. Default is Null. If True, Sheet Tabs selector will be displayed at the bottom of the document.
;                  $bOutlineSymbols     - [optional] a boolean value. Default is Null. If True, the predefined outline symbols will be displayed.
;                  $bCharts             - [optional] a boolean value. Default is Null. If True, Charts are visible in the document.
;                  $bDrawing            - [optional] a boolean value. Default is Null. If True, Drawing Objects are visible in the document.
;                  $bObjects            - [optional] a boolean value. Default is Null. If True, Objects/Graphics are visible in the document.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bHeaders not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bVertScroll not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bHoriScroll not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bSheetTabs not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bOutlineSymbols not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bCharts not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bDrawing not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $bObjects not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Current Controller Object for the Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bHeaders
;                  |                               2 = Error setting $bVertScroll
;                  |                               4 = Error setting $bHoriScroll
;                  |                               8 = Error setting $bSheetTabs
;                  |                               16 = Error setting $bOutlineSymbols
;                  |                               32 = Error setting $bCharts
;                  |                               64 = Error setting $bDrawing
;                  |                               128 = Error setting $bObjects
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the document is visible.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Error setting $bVisible.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. $bVisible successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. Returning current visibility state of the Document, True if visible, false if invisible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call $bVisible with Null to return the current visibility setting.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocVisible(ByRef $oDoc, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($bVisible = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.isVisible())

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.Visible = $bVisible
	$iError = ($oDoc.CurrentController.Frame.ContainerWindow.isVisible() = $bVisible) ? (0) : (1)

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_DocVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocWindowFirstColumn
; Description ...: Set or Retrieve the first visible Column in the Document view.
; Syntax ........: _LOCalc_DocWindowFirstColumn(ByRef $oDoc[, $iColumn = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iColumn             - [optional] an integer value. Default is Null. The column number to set as the first visible column on the page, 0 based.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumn not an Integer, or less than 0, or greater than number of columns contained in the document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iColumn
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. First visible Column was successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning the first visible column number as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  This will fail if there are currently any frozen Columns.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocWindowFirstColumn(ByRef $oDoc, $iColumn = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($iColumn = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.getFirstVisibleColumn())

	If Not __LO_IntIsBetween($iColumn, 0, $oDoc.CurrentController.getActiveSheet().Columns.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.setFirstVisibleColumn($iColumn)

	Return ($oDoc.CurrentController.getFirstVisibleColumn() = $iColumn) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_DocWindowFirstColumn

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocWindowFirstRow
; Description ...: Set or Retrieve the first visible Row in the Document view.
; Syntax ........: _LOCalc_DocWindowFirstRow(ByRef $oDoc[, $iRow = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iRow                - [optional] an integer value. Default is Null. The row number to set as the first visible row on the page, 0 based.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iRow not an Integer, or less than 0, or greater than number of Rows contained in the document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iRow
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. First visible Row was successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning the first visible row number as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  This will fail if there are currently any frozen Rows.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocWindowFirstRow(ByRef $oDoc, $iRow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($iRow = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.getFirstVisibleRow())

	If Not __LO_IntIsBetween($iRow, 0, $oDoc.CurrentController.getActiveSheet().Rows.Count()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.setFirstVisibleRow($iRow)

	Return ($oDoc.CurrentController.getFirstVisibleRow() = $iRow) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOCalc_DocWindowFirstRow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocWindowIsSplit
; Description ...: Query whether the current Document's view is split.
; Syntax ........: _LOCalc_DocWindowIsSplit(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query Document whether the Document view is currently split.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the Document view is currently split.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iX                  - [optional] an integer value. Default is Null. See remarks. The Horizontal (X) position to split the View, in pixels. Set to 0 for no Horizontal split.
;                  $iY                  - [optional] an integer value. Default is Null. See remarks. The Vertical (Y) position to split the View, in pixels. Set to 0 to skip.
;                  $bReturnPixels       - [optional] a boolean value. Default is True. See remarks. If True, return value will be in pixels, Else, return value will be Column Number (X), and Row Number (Y).
; Return values .: Success: 1, 2 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bReturnPixels not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iY not an Integer.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in Pixels, in a 2 Element Array with values in order of function parameters.
;                  @Error 0 @Extended 2 Return Array = Success. All optional parameters were set to Null, returning current settings in Column/Row values, in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To remove the split view, set both $iX and $iY to 0.
;                  $bReturnPixels changes only the return value type, it doesn't change the type of input values to use for $iX and $iY.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve currently visible Range Address.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Sheet Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Cell Range Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning currently visible Range of cells as a Cell Range Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iZoomType           - [optional] an integer value (0 - 4). Default is Null. The Zoom type, See remarks. See constants $LOC_ZOOMTYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iZoom               - [optional] an integer value (20-600). Default is Null. The zoom percentage. Only valid if Zoom type is set to "By Value"
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iZoomType not an Integer, less than 0 or greater than 4. See constants $LOC_ZOOMTYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iZoom not an Integer, less than 20 or greater than 600.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iZoom
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Zoom type always has the value of $LOC_ZOOMTYPE_BY_VALUE(3), when using the other zoom types, the value stays the same, but the zoom level is modified. Consequently, I have not added an error check for the Zoom Type property being correctly set.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocZoom(ByRef $oDoc, $iZoomType = Null, $iZoom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiZoom[0]

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
