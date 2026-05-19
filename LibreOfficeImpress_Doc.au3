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
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Closing, Saving, etc. L.O. Impress documents.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOImpress_DocClose
; _LOImpress_DocConnect
; _LOImpress_DocCreate
; _LOImpress_DocExecuteDispatch
; _LOImpress_DocExport
; _LOImpress_DocGetName
; _LOImpress_DocGetPath
; _LOImpress_DocHasPath
; _LOImpress_DocIsActive
; _LOImpress_DocIsModified
; _LOImpress_DocIsReadOnly
; _LOImpress_DocMaximize
; _LOImpress_DocMinimize
; _LOImpress_DocOpen
; _LOImpress_DocPosAndSize
; _LOImpress_DocRedo
; _LOImpress_DocRedoClear
; _LOImpress_DocRedoCurActionTitle
; _LOImpress_DocRedoGetAllActionTitles
; _LOImpress_DocRedoIsPossible
; _LOImpress_DocSave
; _LOImpress_DocSaveAs
; _LOImpress_DocToFront
; _LOImpress_DocUndo
; _LOImpress_DocUndoActionBegin
; _LOImpress_DocUndoActionEnd
; _LOImpress_DocUndoClear
; _LOImpress_DocUndoCurActionTitle
; _LOImpress_DocUndoGetAllActionTitles
; _LOImpress_DocUndoIsPossible
; _LOImpress_DocUndoReset
; _LOImpress_DocVisible
; _LOImpress_DocZoom
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocClose
; Description ...: Close an existing Impress Document, returning its save path if applicable.
; Syntax ........: _LOImpress_DocClose(ByRef $oDoc[, $bSaveChanges = True[, $sSaveName = ""[, $bDeliverOwnership = True]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bSaveChanges        - [optional] Default is True. If True, saves changes if any were made before closing. See remarks.
;                  $sSaveName           - [optional] Default is "". The file name to save the file as, if the file hasn't been saved before. See Remarks.
;                  $bDeliverOwnership   - [optional] Default is True. If True, deliver ownership of the document Object from the script to LibreOffice, recommended is True.
; Return values .: Success: String
;                  @Error 0 @Extended 1 Return String = Success, Document was successfully closed, and was saved to the returned file Path.
;                  @Error 0 @Extended 2 Return String = Success, Document was successfully closed, document's changes were saved to its existing location.
;                  @Error 0 @Extended 3 Return String = Success, Document was successfully closed, document either had no changes to save, or $bSaveChanges was called with False. If document had a save location, or if document was saved to a location, it is returned, else an empty string is returned.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bSaveChanges not a Boolean.
;                  @Error 1 @Extended 3 = $sSaveName not a String.
;                  @Error 1 @Extended 4 = $bDeliverOwnership not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error while creating Filter Name properties.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Path Conversion to L.O. URL Failed.
;                  @Error 3 @Extended 2 = Error while retrieving Filter Name.
;                  @Error 3 @Extended 3 = Failed to close Document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bSaveChanges is True and the document hasn't been saved yet, the document is saved to the desktop.
;                  If $sSaveName is undefined, it is saved as an .odp document to the desktop, named Year-Month-Day_Hour-Minute-Second.odp. $sSaveName may be a name only without an extension, in which case the file will be saved in .odp format. Or you may define your own format by including an extension, such as "Test.ppt"
; Related .......: _LOImpress_DocOpen, _LOImpress_DocConnect, _LOImpress_DocCreate, _LOImpress_DocSaveAs, _LOImpress_DocSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocClose(ByRef $oDoc, $bSaveChanges = True, $sSaveName = "", $bDeliverOwnership = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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
			$sFilterName = "impress8"
		EndIf

		$sSavePath = _LO_PathConvert($sSavePath & $sSaveName, 1)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If $sFilterName = "" Then $sFilterName = __LOImpress_FilterNameGet($sSavePath)
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
EndFunc   ;==>_LOImpress_DocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocConnect
; Description ...: Connect to an already opened instance of LibreOffice Impress.
; Syntax ........: _LOImpress_DocConnect([$iMode = $LO_DOC_CONNECT_MODE_CURRENT[, $sSearch = ""[, $bCaseless = False]]])
; Parameters ....: $iMode               - [optional] (0-4) Default is $LO_DOC_CONNECT_MODE_CURRENT. The Connect mode. See Constants, $LO_DOC_CONNECT_MODE_* as defined in LibreOffice_Constants.au3.
;                  $sSearch             - [optional] Default is "". The Name, Title or Path of the Document to search for. See remarks.
;                  $bCaseless           - [optional] Default is False. If True, searches are caseless when using $LO_DOC_CONNECT_MODE_SEARCH_* flags.
; Return values .: Success: Object or Array.
;                  @Error 0 @Extended 1 Return Object = Success, The Object for the current, or last active Impress document is returned.
;                  @Error 0 @Extended 1 Return Object = Success, The Object for the found Document with matching Name, Title or Path.
;                  @Error 0 @Extended ? Return Array = Success, An Array of all open LibreOffice Impress Documents. @Extended is set to number of results. See remarks.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $iMode not an Integer, less than 0 or greater than 4. See Constants, $LO_DOC_CONNECT_MODE_* as defined in LibreOffice_Constants.au3.
;                  @Error 1 @Extended 2 = $sSearch not a String.
;                  @Error 1 @Extended 3 = $bCaseless not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error creating ServiceManager object.
;                  @Error 2 @Extended 2 = Error creating Desktop object.
;                  @Error 2 @Extended 3 = Error creating enumeration of open documents.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = No open LibreOffice documents.
;                  @Error 3 @Extended 2 = Failed to retrieve Document Object.
;                  @Error 3 @Extended 3 = Failed to identify Document type.
;                  @Error 3 @Extended 4 = Error converting path to LibreOffice URL.
;                  @Error 3 @Extended 5 = Current Document not a Impress Document.
;                  @Error 3 @Extended 6 = No matches found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only Impress documents are searched or returned using any of the flags.
;                  The value used for $sSearch depends on the flag called in $iMode. It is ignored except for the $LO_DOC_CONNECT_MODE_SEARCH_* flags.
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_TITLE, $sSearch must be the full Title with Office and Component name; e.g: "Test.odp — LibreOffice Impress". This will be the same Title AutoIt would match or return from functions like WinGetTitle.
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_NAME, $sSearch must be the Document's full name, without the extension; e.g: "Test".
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_NAME_WITH_EXT, $sSearch must be the Document's name, with the extension; e.g: "Test.odp". If the Document hasn't been saved, just the name will work, e.g., "Untitled 1".
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_PATH, $sSearch must be the full Path of the document (Name and extension included); e.g: "C:\file\Test.odp."
;                  The Connect All option returns a single columned array. ($aArray[0]), each result is stored in a separate row.
;                  -Row 1 contains the Object for that document. e.g. $aArray[0] = $oDoc
;                  -Row 2 contains the Object for the next document. e.g. $aArray[1] = $oDoc2. And so on.
; Related .......: _LOImpress_DocOpen, _LOImpress_DocClose, _LOImpress_DocCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocConnect($iMode = $LO_DOC_CONNECT_MODE_CURRENT, $sSearch = "", $bCaseless = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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

				If ($iDocType = $LO_DOC_TYPE_IMPRESS) Then
					If (UBound($aoConnectAll) <= $iCount) Then ReDim $aoConnectAll[$iCount + 1]
					$aoConnectAll[$iCount] = $oDoc
					$iCount += 1
				EndIf
				Sleep((IsInt($iCount / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
			WEnd

			Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoConnectAll)

		Case $LO_DOC_CONNECT_MODE_CURRENT
			$oDoc = $oDesktop.currentComponent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$iDocType = _LO_DocGetType($oDoc)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.
			If ($iDocType <> $LO_DOC_TYPE_IMPRESS) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; Not an Impress Doc.

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

				If ($iDocType = $LO_DOC_TYPE_IMPRESS) Then
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
EndFunc   ;==>_LOImpress_DocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocCreate
; Description ...: Open a new LibreOffice Impress Document or Connect to an existing blank, unsaved, writable document.
; Syntax ........: _LOImpress_DocCreate([$bForceNew = True[, $bHidden = False]])
; Parameters ....: $bForceNew           - [optional] Default is True. If True, force opening a new Impress Document instead of checking for a usable blank.
;                  $bHidden             - [optional] Default is False. If True opens the new document invisible or changes the existing document to invisible.
; Return values .: Success: Object
;                  @Error 0 @Extended 1 Return Object = Successfully connected to an existing Document. Returning Document's Object
;                  @Error 0 @Extended 2 Return Object = Successfully created a new document. Returning Document's Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $bForceNew not a Boolean.
;                  @Error 1 @Extended 2 = $bHidden not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Failure Creating Object com.sun.star.ServiceManager.
;                  @Error 2 @Extended 2 = Failure Creating Object com.sun.star.frame.Desktop.
;                  @Error 2 @Extended 3 = Failed to enumerate available documents.
;                  @Error 2 @Extended 4 = Failure Creating New Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Document Object is still returned. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHidden
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DocOpen, _LOImpress_DocClose, _LOImpress_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocCreate($bForceNew = True, $bHidden = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $aArgs[1]
	Local $iError = 0
	Local $oServiceManager, $oDesktop, $oDoc, $oEnumDoc
	Local $sServiceName = "com.sun.star.presentation.PresentationDocument"

	If Not IsBool($bForceNew) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$aArgs[0] = __LO_SetPropertyValue("Hidden", $bHidden)
	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; If not force new, and L.O pages exist then see if there are any blank Impress documents to use.
	If Not $bForceNew And $oDesktop.getComponents.hasElements() Then
		$oEnumDoc = $oDesktop.getComponents.createEnumeration()
		If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService($sServiceName) _
					And Not ($oDoc.hasLocation() And Not $oDoc.isReadOnly()) And Not ($oDoc.isModified()) Then
				$oDoc.CurrentController.Frame.ContainerWindow.Visible = ($bHidden) ? (False) : (True) ; opposite value of $bHidden.
				$iError = ($oDoc.CurrentController.Frame.isHidden() = $bHidden) ? ($iError) : (BitOR($iError, 1))

				Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))
			EndIf
		WEnd
	EndIf

	If Not IsObj($aArgs[0]) Then $iError = BitOR($iError, 1)
	$oDoc = $oDesktop.loadComponentFromURL("private:factory/simpress", "_blank", $iURLFrameCreate, $aArgs)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOImpress_DocCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocExecuteDispatch
; Description ...: Executes a command for a document.
; Syntax ........: _LOImpress_DocExecuteDispatch(ByRef $oDoc, $sDispatch)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sDispatch           - The Dispatch command to execute. See List of commands below.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully executed dispatch command.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sDispatch not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Dispatch is essentially a simulation of the user performing an action, such as pressing Ctrl+A to select all, etc.
;                  Dispatch Commands:
;                  - uno:FullScreen -- Toggles full screen mode.
;                  - uno:ChangeCaseToLower -- Changes all selected text to lower case. Text must be selected with the ViewCursor.
;                  - uno:ChangeCaseToUpper -- Changes all selected text to upper case. Text must be selected with the ViewCursor.
;                  - uno:ChangeCaseRotateCase -- Cycles the Case (Title Case, Sentence case, UPPERCASE, lowercase). Text must be selected with the ViewCursor.
;                  - uno:ChangeCaseToSentenceCase -- Changes the sentence to Sentence case where the ViewCursor is currently positioned or has selected.
;                  - uno:ChangeCaseToTitleCase -- Changes the selected text to Title case. Text must be selected with the ViewCursor.
;                  - uno:ChangeCaseToToggleCase -- Toggles the selected text's case (A becomes a, b becomes B, etc.).Text must be selected with the ViewCursor.
;                  - uno:Delete -- Simulates pressing the Delete key.
;                  - uno:InsertDateFieldFix -- Insert a fixed Date field.
;                  - uno:InsertDateFieldVar -- Insert a variable Date field.
;                  - uno:InsertPageField -- Insert a current Page (slide) field.
;                  - uno:InsertPageTitleField -- Insert a current Page (slide) Title field.
;                  - uno:InsertPagesField -- Insert a total Pages (slides) field.
;                  - uno:InsertPageQuick -- Insert a new page.
;                  - uno:InsertTimeFieldFix -- Insert a fixed Time field.
;                  - uno:InsertTimeFieldVar -- Insert a variable Time field.
;                  - uno:MovePageDown -- Move the currently active slide Down one position,
;                  - uno:MovePageFirst -- Move the currently active slide to the First slide position.
;                  - uno:MovePageLast -- Move the currently active slide to the Last slide position.
;                  - uno:MovePageUp -- Move the currently active slide Up one position.
;                  - uno:Paste -- Pastes the data out of the clipboard. Simulating Ctrl+V.
;                  - uno:PasteUnformatted -- Pastes the data out of the clipboard unformatted.
;                  - uno:PasteSpecial -- Simulates pasting with Ctrl+Shift+V, opens a dialog for selecting paste format.
;                  - uno:Copy -- Simulates Ctrl+C, copies selected data to the clipboard. Text must be selected with the ViewCursor.
;                  - uno:Cut -- Simulates Ctrl+X, cuts selected data, placing it into the clipboard. Text must be selected with the ViewCursor.
;                  - uno:SelectAll -- Simulates Ctrl+A being pressed at the ViewCursor location.
;                  - uno:Zoom50Percent -- Set the zoom level to 50%.
;                  - uno:Zoom75Percent -- Set the zoom level to 75%.
;                  - uno:Zoom100Percent -- Set the zoom level to 100%.
;                  - uno:Zoom150Percent -- Set the zoom level to 150%.
;                  - uno:Zoom200Percent -- Set the zoom level to 200%.
;                  - uno:ZoomMinus -- Decreases the zoom value to the next increment down.
;                  - uno:ZoomPlus -- Increases the zoom value to the next increment up.
;                  - uno:ZoomPageWidth -- Set zoom to fit page width.
;                  - uno:ZoomPage -- Set zoom to fit page.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocExecuteDispatch(ByRef $oDoc, $sDispatch)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aArray[0]
	Local $oServiceManager, $oDispatcher

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sDispatch) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oDispatcher.executeDispatch($oDoc.CurrentController(), "." & $sDispatch, "", 0, $aArray)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DocExecuteDispatch

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocExport
; Description ...: Export a Document with the specified file name to the path specified, with any parameters used.
; Syntax ........: _LOImpress_DocExport(ByRef $oDoc, $sFilePath[, $bSamePath = False[, $sFilterName = ""[, $bOverwrite = Null[, $sPassword = Null]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sFilePath           - Full path to save the document to, including Filename and extension. See Remarks.
;                  $bSamePath           - [optional] Default is False. If True, uses the path of the current document to export to. See Remarks
;                  $sFilterName         - [optional] Default is "". Filter name. If called with "" (blank string), Filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .odp extension is used instead, with the filter name of "impress8".
;                  $bOverwrite          - [optional] Default is Null. If True, file will be overwritten.
;                  $sPassword           - [optional] Default is Null. Password String to set for the document. (Not all file formats can have a Password set). "" (blank string) or Null = No Password.
; Return values .: Success: String
;                  @Error 0 @Extended 0 Return String = Success. Returning save path for exported document.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sFilePath not a String.
;                  @Error 1 @Extended 3 = $bSamePath not a Boolean.
;                  @Error 1 @Extended 4 = $sFilterName not a String.
;                  @Error 1 @Extended 5 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 6 = $sPassword not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error creating Filter Name Property
;                  @Error 2 @Extended 2 = Error creating Overwrite Property
;                  @Error 2 @Extended 3 = Error creating Password Property
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error Converting Path to/from L.O. URL
;                  @Error 3 @Extended 2 = Document has no save path, and $bSamePath is called with True.
;                  @Error 3 @Extended 3 = Error retrieving Filter Name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Does not alter the original save path (if there was one), saves a copy of the document to the new path, in the new file format if one is chosen.
;                  If $bSamePath is called with True, the same save path as the current document is used. You must still fill in "$sFilePath" with the desired File Name and new extension, but you do not need to enter the file path.
; Related .......: _LOImpress_DocSave, _LOImpress_DocSaveAs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocExport(ByRef $oDoc, $sFilePath, $bSamePath = False, $sFilterName = "", $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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

	If ($sFilterName = "") Or ($sFilterName = " ") Then $sFilterName = __LOImpress_FilterNameGet($sFilePath, True)
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
EndFunc   ;==>_LOImpress_DocExport

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocGetName
; Description ...: Retrieve the document's name.
; Syntax ........: _LOImpress_DocGetName(ByRef $oDoc[, $bReturnFull = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bReturnFull         - [optional] Default is False. If True, the full window title is returned, such as is used by AutoIt window related functions.
; Return values .: Success: String
;                  @Error 0 @Extended 0 Return String = Success. Returning the document's Name as a String. See remarks.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bReturnFull not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve Document's name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bReturnFull is True, the return value will be like: "<Impress Doc name>.<extension> — LibreOffice Impress" e.g. "Testing.odp — LibreOffice Impress".
;                  Else the return value will be like: "<Impress Doc name>.<extension>", e.g. "Testing.odp"
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocGetName(ByRef $oDoc, $bReturnFull = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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
EndFunc   ;==>_LOImpress_DocGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocGetPath
; Description ...: Returns a Document's current save path.
; Syntax ........: _LOImpress_DocGetPath(ByRef $oDoc[, $bReturnLibreURL = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bReturnLibreURL     - [optional] Default is False. If True, returns a path in LibreOffice URL format, else False returns a regular Windows path.
; Return values .: Success: String
;                  @Error 0 @Extended 0 Return String = Success. Returning the document's save path as a String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bReturnLibreURL not a Boolean.
;                  @Error 1 @Extended 3 = Document has no save path.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error converting Libre URL to Computer path format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LO_PathConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocGetPath(ByRef $oDoc, $bReturnLibreURL = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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
EndFunc   ;==>_LOImpress_DocGetPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocHasPath
; Description ...: Returns whether a document has been saved to a location already or not.
; Syntax ........: _LOImpress_DocHasPath(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Boolean
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the document has a save location. Else False.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query Document whether it has a path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocHasPath(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bHasPath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bHasPath = $oDoc.hasLocation()
	If Not IsBool($bHasPath) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bHasPath)
EndFunc   ;==>_LOImpress_DocHasPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocIsActive
; Description ...: Tests if called document is the active document of other Libre windows.
; Syntax ........: _LOImpress_DocIsActive(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Boolean
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if document is the currently active Libre window. See remarks.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query Document whether it is active.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This does NOT test if the document is the current active window in Windows, it only tests if the document is the current active document among other LibreOffice documents.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocIsActive(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsActive

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsActive = $oDoc.CurrentController.Frame.isActive()
	If Not IsBool($bIsActive) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsActive)
EndFunc   ;==>_LOImpress_DocIsActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocIsModified
; Description ...: Test whether the document has been modified since being created or since the last save.
; Syntax ........: _LOImpress_DocIsModified(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Boolean
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the document has been modified since last being saved.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query Document whether it has been modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocIsModified(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsMod

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsMod = $oDoc.isModified()
	If Not IsBool($bIsMod) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsMod)
EndFunc   ;==>_LOImpress_DocIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocIsReadOnly
; Description ...: Tests whether a document is opened in Read Only mode.
; Syntax ........: _LOImpress_DocIsReadOnly(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Boolean
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True is document is currently Read Only, else False.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether Document is Read-Only.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only documents that have been saved to a location, will ever be "ReadOnly".
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocIsReadOnly(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsReadOnly

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsReadOnly = $oDoc.isReadOnly()
	If Not IsBool($bIsReadOnly) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsReadOnly)
EndFunc   ;==>_LOImpress_DocIsReadOnly

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocMaximize
; Description ...: Maximize or restore a document.
; Syntax ........: _LOImpress_DocMaximize(ByRef $oDoc[, $bMaximize = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bMaximize           - [optional] Default is Null. If True, document window is maximized, else if False, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean.
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully maximized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMaximize called with Null, returning boolean indicating if Document is currently maximized (True) or not (False).
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bMaximize not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether Document is Maximized.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bMaximize
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocMaximize(ByRef $oDoc, $bMaximize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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
EndFunc   ;==>_LOImpress_DocMaximize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocMinimize
; Description ...: Minimize or restore a document.
; Syntax ........: _LOImpress_DocMinimize(ByRef $oDoc[, $bMinimize = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bMinimize           - [optional] Default is Null. If True, document window is minimized, else if False, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully minimized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMinimize called with Null, returning boolean indicating if Document is currently minimized (True) or not (False).
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bMinimize not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether Document is Minimized.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bMinimize
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocMinimize(ByRef $oDoc, $bMinimize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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
EndFunc   ;==>_LOImpress_DocMinimize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocOpen
; Description ...: Open an existing Impress Document, returning its object identifier.
; Syntax ........: _LOImpress_DocOpen($sFilePath[, $bConnectIfOpen = True[, $bHidden = Null[, $bReadOnly = Null[, $sPassword = Null[, $bLoadAsTemplate = Null[, $sFilterName = Null]]]]]])
; Parameters ....: $sFilePath           - Full path and filename of the file to be opened.
;                  $bConnectIfOpen      - [optional] Default is True(Connect). Whether to connect to the requested document if it is already open. See remarks.
;                  $bHidden             - [optional] Default is Null. If True, opens the document invisibly.
;                  $bReadOnly           - [optional] Default is Null. If True, opens the document as read-only.
;                  $sPassword           - [optional] Default is Null. The password that was used to read-protect the document, if any.
;                  $bLoadAsTemplate     - [optional] Default is Null. If True, opens the document as a Template, i.e. an untitled copy of the specified document is made instead of modifying the original document.
;                  $sFilterName         - [optional] Default is Null. Name of a LibreOffice filter to use to load the specified document. LibreOffice automatically selects which to use by default.
; Return values .: Success: Object.
;                  @Error 0 @Extended 1 Return Object = Successfully connected to requested Document without requested parameters. Returning Document's Object.
;                  @Error 0 @Extended 2 Return Object = Successfully opened requested Document with requested parameters. Returning Document's Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $sFilePath not string, or file not found.
;                  @Error 1 @Extended 2 = Error converting file path to URL path.
;                  @Error 1 @Extended 3 = $bConnectIfOpen not a Boolean.
;                  @Error 1 @Extended 4 = $bHidden not a Boolean.
;                  @Error 1 @Extended 5 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 6 = $sPassword not a string.
;                  @Error 1 @Extended 7 = $bLoadAsTemplate not a Boolean.
;                  @Error 1 @Extended 8 = $sFilterName not a string.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Failed to create ServiceManager Object
;                  @Error 2 @Extended 2 = Failed to create Desktop Object
;                  @Error 2 @Extended 3 = Failed opening or connecting to document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHidden
;                  |                               2 = Error setting $bReadOnly
;                  |                               4 = Error setting $sPassword
;                  |                               8 = Error setting $bLoadAsTemplate
;                  |                               16 = Error setting $sFilterName
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Any parameters (Hidden, template etc.,) will not be applied when connecting to a document.
; Related .......: _LOImpress_DocCreate, _LOImpress_DocClose, _LOImpress_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocOpen($sFilePath, $bConnectIfOpen = True, $bHidden = Null, $bReadOnly = Null, $sPassword = Null, $bLoadAsTemplate = Null, $sFilterName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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

	If $bConnectIfOpen Then $oDoc = _LOImpress_DocConnect($LO_DOC_CONNECT_MODE_SEARCH_PATH, $sFilePath, True)
	If IsObj($oDoc) Then Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))

	$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOImpress_DocOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocPosAndSize
; Description ...: Reposition and resize a document window.
; Syntax ........: _LOImpress_DocPosAndSize(ByRef $oDoc[, $iX = Null[, $iY = Null[, $iWidth = Null[, $iHeight = Null]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iX                  - [optional] Default is Null. The X coordinate of the window.
;                  $iY                  - [optional] Default is Null. The Y coordinate of the window.
;                  $iWidth              - [optional] Default is Null. The width of the window, in pixels(?).
;                  $iHeight             - [optional] Default is Null. The height of the window, in pixels(?).
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $iX not an Integer.
;                  @Error 1 @Extended 3 = $iY not an Integer.
;                  @Error 1 @Extended 4 = $iWidth not an Integer.
;                  @Error 1 @Extended 5 = $iHeight not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error retrieving Position and Size Structure Object.
;                  @Error 3 @Extended 2 = Error retrieving Position and Size Structure Object for error checking.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iX
;                  |                               2 = Error setting $iY
;                  |                               4 = Error setting $iWidth
;                  |                               8 = Error setting $iHeight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: X & Y, on my computer at least, seem to go no lower than 8(X) and 30(Y), if you enter lower than this, it will cause a "property setting Error".
;                  If you want more accurate functionality, use the "WinMove" AutoIt function.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocPosAndSize(ByRef $oDoc, $iX = Null, $iY = Null, $iWidth = Null, $iHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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
EndFunc   ;==>_LOImpress_DocPosAndSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocRedo
; Description ...: Perform one Redo action for a document.
; Syntax ........: _LOImpress_DocRedo(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Successfully performed a redo action.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Document does not have a redo action to perform.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DocUndo, _LOImpress_DocRedoIsPossible, _LOImpress_DocRedoGetAllActionTitles, _LOImpress_DocRedoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocRedo(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isRedoPossible()) Then
		$oDoc.UndoManager.Redo()

		Return SetError($__LO_STATUS_SUCCESS, 1, 0)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOImpress_DocRedo

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocRedoClear
; Description ...: Clear all Redo Actions in the Redo Action List.
; Syntax ........: _LOImpress_DocRedoClear(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared all Redo Actions.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This will silently fail if there are any _LOImpress_DocUndoActionBegin still active.
; Related .......: _LOImpress_DocUndoClear, _LOImpress_DocUndoReset
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocRedoClear(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.clearRedo()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DocRedoClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocRedoCurActionTitle
; Description ...: Retrieve the current Redo action Title.
; Syntax ........: _LOImpress_DocRedoCurActionTitle(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: String
;                  @Error 0 @Extended 0 Return String = Returning the current available redo action title as a String. Will be an empty String if no action is available.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve current Redo Action.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DocRedo, _LOImpress_DocRedoGetAllActionTitles
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocRedoCurActionTitle(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sRedoAction

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sRedoAction = $oDoc.UndoManager.getCurrentRedoActionTitle()
	If ($sRedoAction = Null) Then $sRedoAction = ""
	If Not IsString($sRedoAction) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sRedoAction)
EndFunc   ;==>_LOImpress_DocRedoCurActionTitle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocRedoGetAllActionTitles
; Description ...: Retrieve all available Redo action Titles.
; Syntax ........: _LOImpress_DocRedoGetAllActionTitles(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Array.
;                  @Error 0 @Extended ? Return Array = Returning all available redo action Titles in an array of Strings. @Extended set to number of results.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve an array of Redo action titles.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DocRedo, _LOImpress_DocRedoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocRedoGetAllActionTitles(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asTitles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asTitles = $oDoc.UndoManager.getAllRedoActionTitles()
	If Not IsArray($asTitles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asTitles), $asTitles)
EndFunc   ;==>_LOImpress_DocRedoGetAllActionTitles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocRedoIsPossible
; Description ...: Test whether a Redo action is available to perform for a document.
; Syntax ........: _LOImpress_DocRedoIsPossible(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Boolean
;                  @Error 0 @Extended 0 Return Boolean = If the document has a redo action to perform, True is returned, else False.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether a Redo is possible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DocRedo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocRedoIsPossible(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsRedoPoss

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsRedoPoss = $oDoc.UndoManager.isRedoPossible()
	If Not IsBool($bIsRedoPoss) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $bIsRedoPoss)
EndFunc   ;==>_LOImpress_DocRedoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocSave
; Description ...: Save any changes made to a Document.
; Syntax ........: _LOImpress_DocSave(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Document Successfully saved.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Document is Read Only or Document has no save location, try SaveAs.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DocExport, _LOImpress_DocSaveAs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocSave(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDoc.hasLocation Or $oDoc.isReadOnly Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oDoc.store()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DocSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocSaveAs
; Description ...: Save a Document with the specified file name to the path specified with any parameters called.
; Syntax ........: _LOImpress_DocSaveAs(ByRef $oDoc, $sFilePath[, $sFilterName = ""[, $bOverwrite = Null[, $sPassword = Null]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sFilePath           - Full path to save the document to, including Filename and extension.
;                  $sFilterName         - [optional] Default is "". The filter name. Calling "" (blank string), means the filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .odp extension is used instead, with the filter name of "impress8".
;                  $bOverwrite          - [optional] Default is Null. If True, the existing file will be overwritten.
;                  $sPassword           - [optional] Default is Null. Sets a password for the document. (Not all file formats can have a Password set). Null or "" (blank string) = No Password.
; Return values .: Success: String
;                  @Error 0 @Extended 0 Return String = Successfully Saved the document. Returning document save path.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sFilePath not a String.
;                  @Error 1 @Extended 3 = $sFilterName not a String.
;                  @Error 1 @Extended 4 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 5 = $sPassword not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error creating Filter Name Property
;                  @Error 2 @Extended 2 = Error creating Overwrite Property
;                  @Error 2 @Extended 3 = Error creating Password Property
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error Converting Path to/from L.O. URL
;                  @Error 3 @Extended 2 = Error retrieving Filter Name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Alters original save path (if there was one) to the new path.
; Related .......: _LOImpress_DocExport, _LOImpress_DocSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocSaveAs(ByRef $oDoc, $sFilePath, $sFilterName = "", $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aProperties[1]
	Local $sSavePath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFilterName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$sFilePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFilterName = "") Or ($sFilterName = " ") Then $sFilterName = __LOImpress_FilterNameGet($sFilePath)
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
EndFunc   ;==>_LOImpress_DocSaveAs

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocToFront
; Description ...: Bring the called document to the front of the other windows.
; Syntax ........: _LOImpress_DocToFront(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Window was successfully brought to the front of the open windows.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If minimized, the document is restored and brought to the front of the visible pages. Generally only brings the document to the front of other LibreOffice windows.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocToFront(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.toFront()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DocToFront

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocUndo
; Description ...: Perform one Undo action for a document.
; Syntax ........: _LOImpress_DocUndo(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Successfully performed an undo action.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Document does not have an undo action to perform.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DocUndoIsPossible, _LOImpress_DocUndoGetAllActionTitles, _LOImpress_DocUndoCurActionTitle, _LOImpress_DocRedo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocUndo(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isUndoPossible()) Then
		$oDoc.UndoManager.Undo()

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOImpress_DocUndo

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocUndoActionBegin
; Description ...: Begin an Undo Action group.
; Syntax ........: _LOImpress_DocUndoActionBegin(ByRef $oDoc[, $sName = "AU3LO-Automation"])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $sName               - [optional] Default is "AU3LO-Automation". The name of the Undo Action to display in the UI when completed.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully began an Undo Action group recording.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sName not a String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This begins an Undo Action Group, any functions and actions done after this function is called will be grouped together, and if undone, all actions will be undone together at once.
;                  _LOImpress_DocUndoActionEnd must be called after this function before this undo group will become available in the Undo Action list.
;                  _LOImpress_DocUndoActionBegin can be nested, e.g. call this function multiple times without ending the first undo action, but only the last group that is ended with _LOImpress_DocUndoActionEnd will appear.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocUndoActionBegin(ByRef $oDoc, $sName = "AU3LO-Automation")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.UndoManager.enterUndoContext($sName)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DocUndoActionBegin

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocUndoActionEnd
; Description ...: End the last started Undo Action Group.
; Syntax ........: _LOImpress_DocUndoActionEnd(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully ended the last Undo Action group recording.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This stops the grouping of actions into the last created Undo Action Group.
; Related .......: _LOImpress_DocUndoActionBegin
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocUndoActionEnd(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.leaveUndoContext()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DocUndoActionEnd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocUndoClear
; Description ...: Clear all Undo and Redo Actions in the Undo/Redo Action List.
; Syntax ........: _LOImpress_DocUndoClear(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared all Undo and Redo Actions.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This will silently fail if there are any _LOImpress_DocUndoActionBegin still active.
; Related .......: _LOImpress_DocRedoClear, _LOImpress_DocUndoReset
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocUndoClear(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.clear()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DocUndoClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocUndoCurActionTitle
; Description ...: Retrieve the current Undo action Title.
; Syntax ........: _LOImpress_DocUndoCurActionTitle(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: String
;                  @Error 0 @Extended 0 Return String = Returning the current available Undo action title as a String. Will be an empty String if no action is available.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve current Undo Action.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DocUndo, _LOImpress_DocUndoGetAllActionTitles
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocUndoCurActionTitle(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sUndoAction

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sUndoAction = $oDoc.UndoManager.getCurrentUndoActionTitle()
	If ($sUndoAction = Null) Then $sUndoAction = ""
	If Not IsString($sUndoAction) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sUndoAction)
EndFunc   ;==>_LOImpress_DocUndoCurActionTitle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocUndoGetAllActionTitles
; Description ...: Retrieve all available Undo action Titles.
; Syntax ........: _LOImpress_DocUndoGetAllActionTitles(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Array.
;                  @Error 0 @Extended ? Return Array = Returning all available undo action Titles in an array of Strings. @Extended set to number of results.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve an array of Undo action titles.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DocUndo, _LOImpress_DocUndoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocUndoGetAllActionTitles(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asTitles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asTitles = $oDoc.UndoManager.getAllUndoActionTitles()
	If Not IsArray($asTitles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asTitles), $asTitles)
EndFunc   ;==>_LOImpress_DocUndoGetAllActionTitles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocUndoIsPossible
; Description ...: Test whether a Undo action is available to perform for a document.
; Syntax ........: _LOImpress_DocUndoIsPossible(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Boolean
;                  @Error 0 @Extended 0 Return Boolean = If the document has an undo action to perform, True is returned, else False.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether an Undo is possible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_DocUndo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocUndoIsPossible(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsUndoPoss

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsUndoPoss = $oDoc.UndoManager.isUndoPossible()
	If Not IsBool($bIsUndoPoss) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $bIsUndoPoss)
EndFunc   ;==>_LOImpress_DocUndoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocUndoReset
; Description ...: Reset the Undo Manager.
; Syntax ........: _LOImpress_DocUndoReset(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully reset the undo manager.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling this function does the following: remove all locks from the undo manager; closes all open undo group actions, clears all undo actions, clears all redo actions.
; Related .......: _LOImpress_DocRedoClear, _LOImpress_DocUndoClear
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocUndoReset(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.reset()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_DocUndoReset

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocVisible
; Description ...: Set or retrieve the current visibility of a document.
; Syntax ........: _LOImpress_DocVisible(ByRef $oDoc[, $bVisible = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $bVisible            - [optional] Default is Null. If True, the document is visible.
; Return values .: Success: 1 or Boolean.
;                  @Error 0 @Extended 0 Return 1 = Success. $bVisible successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. Returning current visibility state of the Document, True if visible, False if invisible.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bVisible not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether Document is visible.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current visibility setting.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocVisible(ByRef $oDoc, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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
EndFunc   ;==>_LOImpress_DocVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DocZoom
; Description ...: Modify the zoom value for a document.
; Syntax ........: _LOImpress_DocZoom(ByRef $oDoc[, $iZoomType = Null[, $iZoom = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $iZoomType           - [optional] (0-4) Default is Null. The Zoom type, See remarks. See constants $LOI_ZOOMTYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iZoom               - [optional] (20-600) Default is Null. The zoom percentage. Only valid if Zoom type is set to "By Value"
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $iZoomType not an Integer, less than 0 or greater than 4. See constants $LOI_ZOOMTYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error 1 @Extended 3 = $iZoom not an Integer, less than 20 or greater than 600.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iZoom
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Zoom type always has the value of $LOI_ZOOMTYPE_BY_VALUE(3), when using the other zoom types, the value stays the same, but the zoom level is modified. Consequently, I have not added an error check for the Zoom Type property being correctly set.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DocZoom(ByRef $oDoc, $iZoomType = Null, $iZoom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiZoom[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iZoomType, $iZoom) Then
		__LO_ArrayFill($aiZoom, $oDoc.CurrentController.ZoomType(), $oDoc.CurrentController.ZoomValue())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiZoom)
	EndIf

	If ($iZoomType <> Null) Then
		If Not __LO_IntIsBetween($iZoomType, $LOI_ZOOMTYPE_OPTIMAL, $LOI_ZOOMTYPE_PAGE_WIDTH_EXACT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.CurrentController.ZoomType = $iZoomType
	EndIf

	If ($iZoom <> Null) Then
		If Not __LO_IntIsBetween($iZoom, 20, 600) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.CurrentController.ZoomValue = $iZoom
		$iError = ($oDoc.CurrentController.ZoomValue() = $iZoom) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DocZoom
