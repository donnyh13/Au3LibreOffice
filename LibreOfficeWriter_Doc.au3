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
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Closing, Saving, Searching, etc. L.O. Writer documents.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_DocClose
; _LOWriter_DocConnect
; _LOWriter_DocConvertTableToText
; _LOWriter_DocConvertTextToTable
; _LOWriter_DocCreate
; _LOWriter_DocDescription
; _LOWriter_DocExecuteDispatch
; _LOWriter_DocExport
; _LOWriter_DocFindAll
; _LOWriter_DocFindAllInRange
; _LOWriter_DocFindNext
; _LOWriter_DocGenProp
; _LOWriter_DocGenPropCreation
; _LOWriter_DocGenPropModification
; _LOWriter_DocGenPropPrint
; _LOWriter_DocGenPropTemplate
; _LOWriter_DocGetCounts
; _LOWriter_DocGetName
; _LOWriter_DocGetPath
; _LOWriter_DocHasPath
; _LOWriter_DocIsActive
; _LOWriter_DocIsModified
; _LOWriter_DocIsReadOnly
; _LOWriter_DocMaximize
; _LOWriter_DocMinimize
; _LOWriter_DocOpen
; _LOWriter_DocPosAndSize
; _LOWriter_DocPrint
; _LOWriter_DocPrintIncludedSettings
; _LOWriter_DocPrintMiscSettings
; _LOWriter_DocPrintPageSettings
; _LOWriter_DocPrintSizeSettings
; _LOWriter_DocRedo
; _LOWriter_DocRedoClear
; _LOWriter_DocRedoCurActionTitle
; _LOWriter_DocRedoGetAllActionTitles
; _LOWriter_DocRedoIsPossible
; _LOWriter_DocReplaceAll
; _LOWriter_DocReplaceAllInRange
; _LOWriter_DocSave
; _LOWriter_DocSaveAs
; _LOWriter_DocSelection
; _LOWriter_DocToFront
; _LOWriter_DocUndo
; _LOWriter_DocUndoActionBegin
; _LOWriter_DocUndoActionEnd
; _LOWriter_DocUndoClear
; _LOWriter_DocUndoCurActionTitle
; _LOWriter_DocUndoGetAllActionTitles
; _LOWriter_DocUndoIsPossible
; _LOWriter_DocUndoReset
; _LOWriter_DocVisible
; _LOWriter_DocZoom
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocClose
; Description ...: Close an existing Writer Document, returning its save path if applicable.
; Syntax ........: _LOWriter_DocClose(ByRef $oDoc[, $bSaveChanges = True[, $sSaveName = ""[, $bDeliverOwnership = True]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
;                  @Error 3 @Extended 2 = Error while retrieving FilterName.
;                  @Error 3 @Extended 3 = Failed to close Document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bSaveChanges is True and the document hasn't been saved yet, the document is saved to the desktop.
;                  If $sSaveName is undefined, it is saved as an .odt document to the desktop, named Year-Month-Day_Hour-Minute-Second.odt.
;                  $sSaveName may be a name only without an extension, in which case the file will be saved in .odt format. Or you may define your own format by including an extension, such as "Test.docx"
; Related .......: _LOWriter_DocOpen, _LOWriter_DocConnect, _LOWriter_DocCreate, _LOWriter_DocSaveAs, _LOWriter_DocSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocClose(ByRef $oDoc, $bSaveChanges = True, $sSaveName = "", $bDeliverOwnership = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
			$sSaveName = @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "-" & @MIN & "-" & @SEC & ".odt"
			$sFilterName = "writer8"
		EndIf

		$sSavePath = _LO_PathConvert($sSavePath & $sSaveName, 1)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If $sFilterName = "" Then $sFilterName = __LOWriter_FilterNameGet($sSavePath)
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
EndFunc   ;==>_LOWriter_DocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocConnect
; Description ...: Connect to an already opened instance of LibreOffice Writer.
; Syntax ........: _LOWriter_DocConnect([$iMode = $LO_DOC_CONNECT_MODE_CURRENT[, $sSearch = ""[, $bCaseless = False]]])
; Parameters ....: $iMode               - [optional] (0-4) Default is $LO_DOC_CONNECT_MODE_CURRENT. The Connect mode. See Constants, $LO_DOC_CONNECT_MODE_* as defined in LibreOffice_Constants.au3.
;                  $sSearch             - [optional] Default is "". The Name, Title or Path of the Document to search for. See remarks.
;                  $bCaseless           - [optional] Default is False. If True, searches are caseless when using $LO_DOC_CONNECT_MODE_SEARCH_* flags.
; Return values .: Success: Object or Array.
;                  @Error 0 @Extended 1 Return Object = Success, The Object for the current, or last active Writer document is returned.
;                  @Error 0 @Extended 1 Return Object = Success, The Object for the found Document with matching Name, Title or Path.
;                  @Error 0 @Extended ? Return Array = Success, An Array of all open LibreOffice Writer Documents. @Extended is set to number of results. See remarks.
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
;                  @Error 3 @Extended 5 = Current Document not a Writer Document.
;                  @Error 3 @Extended 6 = No matches found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only Writer documents are searched or returned using any of the flags.
;                  The value used for $sSearch depends on the flag called in $iMode. It is ignored except for the $LO_DOC_CONNECT_MODE_SEARCH_* flags.
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_TITLE, $sSearch must be the full Title with Office and Component name; e.g: "Test.odt — LibreOffice Writer". This will be the same Title AutoIt would match or return from functions like WinGetTitle.
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_NAME, $sSearch must be the Document's full name, without the extension; e.g: "Test".
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_NAME_WITH_EXT, $sSearch must be the Document's name, with the extension; e.g: "Test.odt". If the Document hasn't been saved, just the name will work, e.g., "Untitled 1".
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_PATH, $sSearch must be the full Path of the document (Name and extension included); e.g: "C:\file\Test.odt."
;                  The Connect All option returns a single columned array. ($aArray[0]), each result is stored in a separate row.
;                  -Row 1 contains the Object for that document. e.g. $aArray[0] = $oDoc
;                  -Row 2 contains the Object for the next document. e.g. $aArray[1] = $oDoc2. And so on.
; Related .......: _LOWriter_DocOpen, _LOWriter_DocClose, _LOWriter_DocCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocConnect($iMode = $LO_DOC_CONNECT_MODE_CURRENT, $sSearch = "", $bCaseless = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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

				If ($iDocType = $LO_DOC_TYPE_WRITER) Then
					If (UBound($aoConnectAll) <= $iCount) Then ReDim $aoConnectAll[$iCount + 1]
					$aoConnectAll[$iCount] = $oDoc
					$iCount += 1
				EndIf
				Sleep((IsInt($iCount / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
			WEnd

			Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoConnectAll)

		Case $LO_DOC_CONNECT_MODE_CURRENT
			$oDoc = $oDesktop.currentComponent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$iDocType = _LO_DocGetType($oDoc)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.
			If ($iDocType <> $LO_DOC_TYPE_WRITER) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; Not a Writer Doc.

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

				If ($iDocType = $LO_DOC_TYPE_WRITER) Then
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
EndFunc   ;==>_LOWriter_DocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocConvertTableToText
; Description ...: Convert a Table to Text, separated by a delimiter.
; Syntax ........: _LOWriter_DocConvertTableToText(ByRef $oDoc, ByRef $oTable[, $sDelimiter = @TAB])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oTable              - A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $sDelimiter          - [optional] Default is @TAB. A character to separate each column by, such as a Tab character, etc.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Table was successfully converted to text.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $oTable not an Object.
;                  @Error 1 @Extended 3 = $sDelimiter not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 = Failed to create "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve array of CellNames.
;                  @Error 3 @Extended 2 = Failed to create a backup of the ViewCursor's current location.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function temporarily moves the Viewcursor to the Table indicated, and then attempts to restore the ViewCursor to its former position.
;                  This could cause a COM error if the Cursor was presently in the Table.
; Related .......: _LOWriter_DocConvertTextToTable, _LOWriter_TableGetObjByName, _LOWriter_TableGetObjByCursor, _LOWriter_TableCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocConvertTableToText(ByRef $oDoc, ByRef $oTable, $sDelimiter = @TAB)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aArgs[1]
	Local $asCellNames
	Local $oServiceManager, $oDispatcher, $oSelection

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sDelimiter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$aArgs[0] = __LO_SetPropertyValue("Delimiter", $sDelimiter)

	$asCellNames = $oTable.getCellNames()
	If Not IsArray($asCellNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	; Backup the ViewCursor location and selection.
	$oSelection = $oDoc.getCurrentSelection()
	If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	; Select the Table.
	$oDoc.CurrentController.Select($oTable)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:ConvertTableToText", "", 0, $aArgs)

	; Restore the ViewCursor to its previous location.
	$oDoc.CurrentController.Select($oSelection)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocConvertTableToText

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocConvertTextToTable
; Description ...: Convert some selected text into a Table.
; Syntax ........: _LOWriter_DocConvertTextToTable(ByRef $oDoc, ByRef $oCursor[, $sDelimiter = @TAB[, $bHeader = False[, $iRepeatHeaderLines = 0[, $bBorder = False[, $bDontSplitTable = False]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - A Cursor Object returned from any Cursor Object creation or retrieval functions. See Remarks.
;                  $sDelimiter          - [optional] Default is @TAB. A character to the text into each column by, such as a Tab etc.
;                  $bHeader             - [optional] Default is False. If True, Formats the first row of the new table as a heading.
;                  $iRepeatHeaderLines  - [optional] Default is 0. If greater than 0, then Repeats the first n rows as a header.
;                  $bBorder             - [optional] Default is False. If True, Adds a border to the table and the table cells.
;                  $bDontSplitTable     - [optional] Default is False. If True, Does not divide the table across pages.
; Return values .: Success: Object
;                  @Error 0 @Extended 0 Return Object = Success. Text was successfully converted to a Table, returning the new Table's Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $oCursor not an Object.
;                  @Error 1 @Extended 3 = $sDelimiter not a String.
;                  @Error 1 @Extended 4 = $bHeader not a Boolean.
;                  @Error 1 @Extended 5 = $iRepeatHeaderLines not an Integer.
;                  @Error 1 @Extended 6 = $bBorder not a Boolean.
;                  @Error 1 @Extended 7 = $bDontSplitTable not a Boolean.
;                  @Error 1 @Extended 8 = $oCursor is a Table Cursor and is not supported.
;                  @Error 1 @Extended 9 = $oCursor has no data selected.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 = Failed to create "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve TextTables Object.
;                  @Error 3 @Extended 2 = Failed to retrieve array of Table names.
;                  @Error 3 @Extended 3 = Failed to identify $oCursor's cursor type.
;                  @Error 3 @Extended 4 = Failed to backup ViewCursor's position.
;                  @Error 3 @Extended 5 = Failed to retrieve new Table's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function temporarily moves the ViewCursor to and selects the Text, and then attempts to restore the ViewCursor to its former position.
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_CursorParObjCreateList, _LOWriter_DocConvertTableToText
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocConvertTextToTable(ByRef $oDoc, ByRef $oCursor, $sDelimiter = @TAB, $bHeader = False, $iRepeatHeaderLines = 0, $bBorder = False, $bDontSplitTable = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asTables[0]
	Local $atArgs[5]
	Local $oServiceManager, $oDispatcher, $oTables, $oTable, $oSelection
	Local $iCursorType

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sDelimiter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bHeader) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iRepeatHeaderLines) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bBorder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bDontSplitTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$oTables = $oDoc.TextTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	; Store all current Table Names.
	$asTables = $oTables.getElementNames()
	If Not IsArray($asTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	; If Cursor has no data selected, return error.
	If $oCursor.isCollapsed() Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$atArgs[0] = __LO_SetPropertyValue("Delimiter", $sDelimiter)
	$atArgs[1] = __LO_SetPropertyValue("WithHeader", $bHeader)
	$atArgs[2] = __LO_SetPropertyValue("RepeatHeaderLines", $iRepeatHeaderLines)
	$atArgs[3] = __LO_SetPropertyValue("WithBorder", $bBorder)
	$atArgs[4] = __LO_SetPropertyValue("DontSplitTable", $bDontSplitTable)

	If ($iCursorType = $LOW_CURTYPE_TEXT_CURSOR) Or ($iCursorType = $LOW_CURTYPE_PARAGRAPH) Then
		; Backup the ViewCursor location and selection.
		$oSelection = $oDoc.getCurrentSelection()
		If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oDoc.CurrentController.Select($oCursor)

		$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:ConvertTextToTable", "", 0, $atArgs)

		; Restore the ViewCursor to its previous location.
		$oDoc.CurrentController.Select($oSelection)

	Else
		$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:ConvertTextToTable", "", 0, $atArgs)
	EndIf

	; Obtain the newly created table object by comparing the original table names to the new list of tables.
	; If none match, then it is the new one. Return that Table's Object.
	For $i = 0 To $oTables.getCount() - 1
		For $j = 0 To UBound($asTables) - 1
			If ($asTables[$j] = $oTables.getByIndex($i).Name()) Then ExitLoop
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0))) ; Sleep every x cycles.
		Next

		If ($j = UBound($asTables)) Then ; If No matches in the original table names, then set Table Object and exit loop

			$oTable = $oTables.getByIndex($i)
			ExitLoop
		EndIf
	Next

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTable)
EndFunc   ;==>_LOWriter_DocConvertTextToTable

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocCreate
; Description ...: Open a new LibreOffice Writer Document or Connect to an existing blank, unsaved, writable document.
; Syntax ........: _LOWriter_DocCreate([$bForceNew = True[, $bHidden = False]])
; Parameters ....: $bForceNew           - [optional] Default is True. If True, force opening a new Writer Document instead of checking for a usable blank.
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
; Related .......: _LOWriter_DocOpen, _LOWriter_DocClose, _LOWriter_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocCreate($bForceNew = True, $bHidden = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $aArgs[1]
	Local $iError = 0
	Local $oServiceManager, $oDesktop, $oDoc, $oEnumDoc

	If Not IsBool($bForceNew) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$aArgs[0] = __LO_SetPropertyValue("Hidden", $bHidden)
	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; If not force new, and L.O pages exist then see if there are any blank writer documents to use.
	If Not $bForceNew And $oDesktop.getComponents.hasElements() Then
		$oEnumDoc = $oDesktop.getComponents.createEnumeration()
		If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService("com.sun.star.text.TextDocument") And Not ($oDoc.hasLocation() And Not $oDoc.isReadOnly()) And Not ($oDoc.isModified()) Then
				$oDoc.CurrentController.Frame.ContainerWindow.Visible = ($bHidden) ? (False) : (True) ; opposite value of $bHidden.
				$iError = ($oDoc.CurrentController.Frame.isHidden() = $bHidden) ? ($iError) : (BitOR($iError, 1))

				Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))
			EndIf
		WEnd
	EndIf

	If Not IsObj($aArgs[0]) Then $iError = BitOR($iError, 1)
	$oDoc = $oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", $iURLFrameCreate, $aArgs)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOWriter_DocCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocDescription
; Description ...: Set or Retrieve Document Description properties.
; Syntax ........: _LOWriter_DocDescription(ByRef $oDoc[, $sTitle = Null[, $sSubject = Null[, $asKeywords = Null[, $sComments = Null[, $asContributor = Null[, $sCoverage = Null[, $sIdentifier = Null[, $asPublisher = Null[, $asRelation = Null[, $sRights = Null[, $sSource = Null[, $sType = Null]]]]]]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
;                  $sTitle              - [optional] Default is Null. The Document's "Title" Property. See Remarks.
;                  $sSubject            - [optional] Default is Null. The Document's "Subject" Property.
;                  $asKeywords          - [optional] Default is Null. The Document's "Keywords" Property. Input must be a single dimension Array, which will overwrite any keywords previously set. Accepts numbers also. See Remarks.
;                  $sComments           - [optional] Default is Null. The Document's "Comments" Property.
;                  $asContributor       - [optional] Default is Null. The Document's "Contributor" Property. Input must be a single dimension Array, which will overwrite any values previously set. See Remarks. L.O. 24.2+
;                  $sCoverage           - [optional] Default is Null. The Document's "Coverage" Property. L.O. 24.2+
;                  $sIdentifier         - [optional] Default is Null. The Document's "Identifier" Property. L.O. 24.2+
;                  $asPublisher         - [optional] Default is Null. The Document's "Publisher" Property. Input must be a single dimension Array, which will overwrite any values previously set. See Remarks. L.O. 24.2+
;                  $asRelation          - [optional] Default is Null. The Document's "Relation" Property. Input must be a single dimension Array, which will overwrite any values previously set. See Remarks. L.O. 24.2+
;                  $sRights             - [optional] Default is Null. The Document's "Rights" Property. L.O. 24.2+
;                  $sSource             - [optional] Default is Null. The Document's "Source" Property. L.O. 24.2+
;                  $sType               - [optional] Default is Null. The Document's "Type" Property. L.O. 24.2+
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 12 Element Array with values in order of function parameters. If current LibreOffice version is less than 24.2, $asContributor, $sCoverage, $sIdentifier, $asPublisher, $asRelation, $sRights, $sSource, $sType will return a Null value.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sTitle not a String.
;                  @Error 1 @Extended 3 = $sSubject not a String.
;                  @Error 1 @Extended 4 = $asKeywords not an Array.
;                  @Error 1 @Extended 5 = $sComments not a String.
;                  @Error 1 @Extended 6 = $asContributor not an Array.
;                  @Error 1 @Extended 7 = $sCoverage not a String.
;                  @Error 1 @Extended 8 = $sIdentifier not a String.
;                  @Error 1 @Extended 9 = $asPublisher not an Array.
;                  @Error 1 @Extended 10 = $asRelation not an Array.
;                  @Error 1 @Extended 11 = $sRights not a String.
;                  @Error 1 @Extended 12 = $sSource not a String.
;                  @Error 1 @Extended 13 = $sType not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error retrieving Document Properties Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sTitle
;                  |                               2 = Error setting $sSubject
;                  |                               4 = Error setting $asKeywords
;                  |                               8 = Error setting $sComments
;                  |                               16 = Error setting $asContributor
;                  |                               32 = Error setting $sCoverage
;                  |                               64 = Error setting $sIdentifier
;                  |                               128 = Error setting $asPublisher
;                  |                               256 = Error setting $asRelation
;                  |                               512 = Error setting $sRights
;                  |                               1024 = Error setting $sSource
;                  |                               2048 = Error setting $sType
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 = Current LibreOffice version is less than 24.2.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: "Title" is the Title as found in File>Properties, not the Document's Title as set when saving it.
;                  Any array error checking only checks to make sure the input array, and the set Array of values is the same size, it does not check that each element is the same.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocDescription(ByRef $oDoc, $sTitle = Null, $sSubject = Null, $asKeywords = Null, $sComments = Null, $asContributor = Null, $sCoverage = Null, $sIdentifier = Null, $asPublisher = Null, $asRelation = Null, $sRights = Null, $sSource = Null, $sType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocProp
	Local $iError = 0
	Local $avDescription[12]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sTitle, $sSubject, $asKeywords, $sComments, $asContributor, $sCoverage, $sIdentifier, $asPublisher, $asRelation, $sRights, $sSource, $sType) Then
		If __LO_VersionCheck(24.2) Then ; These properties are only available from L.O. 24.2+.
			__LO_ArrayFill($avDescription, $oDocProp.Title(), $oDocProp.Subject(), $oDocProp.Keywords(), $oDocProp.Description(), $oDocProp.Contributor(), $oDocProp.Coverage(), _
					$oDocProp.Identifier(), $oDocProp.Publisher(), $oDocProp.Relation(), $oDocProp.Rights(), $oDocProp.Source(), $oDocProp.Type())

		Else
			__LO_ArrayFill($avDescription, $oDocProp.Title(), $oDocProp.Subject(), $oDocProp.Keywords(), $oDocProp.Description(), Null, Null, Null, Null, Null, _
					Null, Null, Null)
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avDescription)
	EndIf

	If ($sTitle <> Null) Then
		If Not IsString($sTitle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDocProp.Title = $sTitle
		$iError = ($oDocProp.Title() = $sTitle) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sSubject <> Null) Then
		If Not IsString($sSubject) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDocProp.Subject = $sSubject
		$iError = ($oDocProp.Subject() = $sSubject) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($asKeywords <> Null) Then
		If Not IsArray($asKeywords) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDocProp.Keywords = $asKeywords
		$iError = (UBound($oDocProp.Keywords()) = UBound($asKeywords)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($sComments <> Null) Then
		If Not IsString($sComments) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oDocProp.Description = $sComments
		$iError = ($oDocProp.Description() = $sComments) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($asContributor <> Null) Then
		If Not IsArray($asContributor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not __LO_VersionCheck(24.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oDocProp.Contributor = $asContributor
		$iError = (UBound($oDocProp.Contributor()) = UBound($asContributor)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($sCoverage <> Null) Then
		If Not IsString($sCoverage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		If Not __LO_VersionCheck(24.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oDocProp.Coverage = $sCoverage
		$iError = ($oDocProp.Coverage() = $sCoverage) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($sIdentifier <> Null) Then
		If Not IsString($sIdentifier) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		If Not __LO_VersionCheck(24.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oDocProp.Identifier = $sIdentifier
		$iError = ($oDocProp.Identifier() = $sIdentifier) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($asPublisher <> Null) Then
		If Not IsArray($asPublisher) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not __LO_VersionCheck(24.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oDocProp.Publisher = $asPublisher
		$iError = (UBound($oDocProp.Publisher()) = UBound($asPublisher)) ? ($iError) : (BitOR($iError, 128))
	EndIf

	If ($asRelation <> Null) Then
		If Not IsArray($asRelation) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		If Not __LO_VersionCheck(24.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oDocProp.Relation = $asRelation
		$iError = (UBound($oDocProp.Relation()) = UBound($asRelation)) ? ($iError) : (BitOR($iError, 256))
	EndIf

	If ($sRights <> Null) Then
		If Not IsString($sRights) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
		If Not __LO_VersionCheck(24.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oDocProp.Rights = $sRights
		$iError = ($oDocProp.Rights() = $sRights) ? ($iError) : (BitOR($iError, 512))
	EndIf

	If ($sSource <> Null) Then
		If Not IsString($sSource) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)
		If Not __LO_VersionCheck(24.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oDocProp.Source = $sSource
		$iError = ($oDocProp.Source() = $sSource) ? ($iError) : (BitOR($iError, 1024))
	EndIf

	If ($sType <> Null) Then
		If Not IsString($sType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)
		If Not __LO_VersionCheck(24.2) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oDocProp.Type = $sType
		$iError = ($oDocProp.Type() = $sType) ? ($iError) : (BitOR($iError, 2048))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocDescription

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocExecuteDispatch
; Description ...: Executes a command for a document.
; Syntax ........: _LOWriter_DocExecuteDispatch(ByRef $oDoc, $sDispatch)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
;                  - uno:ChangeCaseToSentenceCase -- Changes the sentence to Sentence case where the Viewcursor is currently positioned or has selected.
;                  - uno:ChangeCaseToTitleCase -- Changes the selected text to Title case. Text must be selected with the ViewCursor.
;                  - uno:ChangeCaseToToggleCase -- Toggles the selected text's case (A becomes a, b becomes B, etc.).Text must be selected with the ViewCursor.
;                  - uno:UpdateAll -- Causes all non fixed Fields, Links, Indexes, Charts etc., to be updated.
;                  - uno:UpdateFields -- Causes all Fields to be updated.
;                  - uno:UpdateAllIndexes -- Causes all Indexes to be updated.
;                  - uno:UpdateAllLinks -- Causes all Links to be updated.
;                  - uno:UpdateCharts -- Causes all Charts to be updated.
;                  - uno:Repaginate -- Update Page Formatting.
;                  - uno:ResetAttributes -- Removes all direct formatting from the selected text. Text must be selected with the ViewCursor.
;                  - uno:SwBackspace -- Simulates pressing the Backspace key.
;                  - uno:Delete -- Simulates pressing the Delete key.
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
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocExecuteDispatch(ByRef $oDoc, $sDispatch)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
EndFunc   ;==>_LOWriter_DocExecuteDispatch

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocExport
; Description ...: Export a Document with the specified file name to the path specified, with any parameters used.
; Syntax ........: _LOWriter_DocExport(ByRef $oDoc, $sFilePath[, $bSamePath = False[, $sFilterName = ""[, $bOverwrite = Null[, $sPassword = Null]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFilePath           - Full path to save the document to, including Filename and extension. See Remarks.
;                  $bSamePath           - [optional] Default is False. If True, uses the path of the current document to export to. See Remarks
;                  $sFilterName         - [optional] Default is "". Filter name. If called with "" (blank string), Filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .odt extension is used instead, with the filter name of "writer8".
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
;                  @Error 2 @Extended 1 = Error creating FilterName Property.
;                  @Error 2 @Extended 2 = Error creating Overwrite Property.
;                  @Error 2 @Extended 3 = Error creating Password Property.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error Converting Path to/from L.O. URL
;                  @Error 3 @Extended 2 = Document has no save path, and $bSamePath is called with True.
;                  @Error 3 @Extended 3 = Error retrieving FilterName.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Does not alter the original save path (if there was one), saves a copy of the document to the new path, in the new file format if one is chosen.
;                  If $bSamePath is called with True, the same save path as the current document is used. You must still fill in "sFilePath" with the desired File Name and new extension, but you do not need to enter the file path.
; Related .......: _LOWriter_DocSave, _LOWriter_DocSaveAs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocExport(ByRef $oDoc, $sFilePath, $bSamePath = False, $sFilterName = "", $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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

	If ($sFilterName = "") Or ($sFilterName = " ") Then $sFilterName = __LOWriter_FilterNameGet($sFilePath, True)
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
EndFunc   ;==>_LOWriter_DocExport

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocFindAll
; Description ...: Find all matches contained in a document of a Specified Search String.
; Syntax ........: _LOWriter_DocFindAll(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString[, $atFindFormat = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - A String of text or a regular expression to search for.
;                  $atFindFormat        - [optional] Default is Null. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended ? Return Array = Success. Search was Successful, returning 1 dimensional array containing the objects to each match, @Extended is set to the number of matches.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 = $oSrchDescriptObject not a Search Descriptor Object.
;                  @Error 1 @Extended 4 = $sSearchString not a String.
;                  @Error 1 @Extended 5 = $atFindFormat not an Array.
;                  @Error 1 @Extended 6 = $atFindFormat does not contain an Object in the first Element.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Search did not return an Object, something went wrong.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Objects returned can be used in any of the functions accepting a Paragraph or Cursor Object etc., to modify their properties or even the text itself.
; Related .......: _LOWriter_SearchDescriptorCreate, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange, _LOWriter_FindFormatModifyAlignment, _LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation, _LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak, _LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace, _LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout, _LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocFindAll(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, $atFindFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResults
	Local $aoResults[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($atFindFormat <> Null) And Not IsArray($atFindFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($atFindFormat <> Null) And (UBound($atFindFormat) > 0) And Not IsObj($atFindFormat[0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If IsArray($atFindFormat) Then $oSrchDescript.setSearchAttributes($atFindFormat)
	$oSrchDescript.SearchString = $sSearchString

	$oResults = $oDoc.findAll($oSrchDescript)
	If Not IsObj($oResults) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oResults.getCount() > 0) Then
		ReDim $aoResults[$oResults.getCount]
		For $i = 0 To $oResults.getCount() - 1
			$aoResults[$i] = $oResults.getByIndex($i)
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoResults), $aoResults)
EndFunc   ;==>_LOWriter_DocFindAll

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocFindAllInRange
; Description ...: Find all occurrences of a Search String in a Document in a specific selection.
; Syntax ........: _LOWriter_DocFindAllInRange(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, ByRef $oRange[, $atFindFormat = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - A String of text or a regular expression to search for.
;                  $oRange              - A Range, such as a cursor with Data selected, to perform the search within.
;                  $atFindFormat        - [optional] Default is Null. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
; Return values .: Success: 1 or Array..
;                  @Error 0 @Extended ? Return Array = Success. Search was Successful, returning 1 dimensional array containing the objects for each match, @Extended is set to the number of matches.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 = $sSearchString not a String.
;                  @Error 1 @Extended 5 = $oRange not an Object.
;                  @Error 1 @Extended 6 = $oRange has no data selected.
;                  @Error 1 @Extended 7 = $atFindFormat not an Array.
;                  @Error 1 @Extended 8 = First element in $atFindFormat not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Search did not return an Object, something went wrong.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_SearchDescriptorCreate, _LOWriter_DocFindAll, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange, _LOWriter_FindFormatModifyAlignment, _LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation, _LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak, _LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace, _LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout, _LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocFindAllInRange(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, ByRef $oRange, $atFindFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResults, $oResult, $oRangeRegion, $oResultRegion, $oText, $oRangeAnchor
	Local $aoResults[0]
	Local $iCount = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($oRange.IsCollapsed()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($atFindFormat <> Null) And Not IsArray($atFindFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If ($atFindFormat <> Null) And (UBound($atFindFormat) > 0) And Not IsObj($atFindFormat[0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	If IsArray($atFindFormat) Then $oSrchDescript.setSearchAttributes($atFindFormat)

	$oSrchDescript.SearchString = $sSearchString

	If $oRange.Text.supportsService("com.sun.star.text.TextFrame") Then
		$oRangeAnchor = $oRange.TextFrame.getAnchor() ; If Range is in a TextFrame, convert its position to a range in the document

	ElseIf $oRange.Text.supportsService("com.sun.star.text.Footnote") Or $oRange.Text.supportsService("com.sun.star.text.Endnote") Then
		$oRangeAnchor = $oRange.Text.Anchor()
	EndIf

	$oResults = $oDoc.findAll($oSrchDescript)
	If Not IsObj($oResults) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oResults.getCount() > 0) Then
		ReDim $aoResults[$oResults.getCount]

		For $i = 0 To $oResults.getCount() - 1
			$oText = $oDoc.Text()
			$oResult = $oResults.getByIndex($i)
			$oResultRegion = $oResult
			$oRangeRegion = $oRange

			If $oResult.Text.supportsService("com.sun.star.text.TextFrame") Then
				$oResultRegion = $oResult.TextFrame.getAnchor() ; If Result is in a TextFrame, convert its position to a range in the document

			ElseIf $oResult.Text.supportsService("com.sun.star.text.Footnote") Or $oResult.Text.supportsService("com.sun.star.text.Endnote") Then
				$oResultRegion = $oResult.Text.Anchor()
			EndIf

			If $oRange.Text.supportsService("com.sun.star.text.TextFrame") And $oResult.Text.supportsService("com.sun.star.text.TextFrame") Then
				If ($oDoc.Text.compareRegionEnds($oResultRegion, $oRangeAnchor) = 0) Then ;  If both Range and Result are in a Text Frame, test if they are in the same one.
					$oResultRegion = $oResult ; If They are, then compare the regions of that text frame.
					$oRangeRegion = $oRangeAnchor
					$oText = $oRange.Text() ; Must use the corresponding Text Object for that TextFrame as Region Compare can only compare regions contained in the same Text Object region.
				EndIf

			ElseIf $oResult.Text.supportsService("com.sun.star.text.Footnote") Or $oResult.Text.supportsService("com.sun.star.text.Endnote") And _
					$oRange.Text.supportsService("com.sun.star.text.Footnote") Or $oRange.Text.supportsService("com.sun.star.text.Endnote") Then
				If ($oDoc.Text.compareRegionEnds($oResultRegion, $oRangeAnchor) = 0) Then ;  If both Range and Result are in a Text Frame, test if they are in the same one.
					$oResultRegion = $oResult ; If They are, then compare the regions of that text frame.
					$oRangeRegion = $oRangeAnchor
					$oText = $oRange.Text() ; Must use the corresponding Text Object for that Foot/Endnote as Region Compare can only compare regions contained in the same Text Object region.
				EndIf
			EndIf

			If ($oText.compareRegionEnds($oResultRegion, $oRangeRegion) >= 0) And ($oText.compareRegionStarts($oRangeRegion, $oResultRegion) >= 0) Then
				$aoResults[$iCount] = $oResult
				$iCount += 1

			Else
				$oResult = Null
			EndIf

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
		ReDim $aoResults[$iCount]
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoResults), $aoResults)
EndFunc   ;==>_LOWriter_DocFindAllInRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocFindNext
; Description ...: Find a Search String in a Document once or one at a time.
; Syntax ........: _LOWriter_DocFindNext(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString[, $atFindFormat = Null[, $oRange = Null[, $oLastFind = Null[, $bExhaustive = False]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - A String of text or a regular expression to search for.
;                  $atFindFormat        - [optional] Default is Null. Call with Null to skip. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
;                  $oRange              - [optional] Default is Null. A Range, such as a cursor with Data selected, to perform the search within. If Null, the entire document is searched.
;                  $oLastFind           - [optional] Default is Null. The last returned Object by a previous call to this function to begin the search from, if called with Null, the search begins at the start of the Document or selection, depending on if a Range is provided.
;                  $bExhaustive         - [optional] Default is False. If True, tests whether every result found in a document is contained in the selection or not. See remarks.
; Return values .: Success: Object or 1.
;                  @Error 0 @Extended 0 Return 1 = Success. Search was successful but found no matches.
;                  @Error 0 @Extended 1 Return Object = Success. Search was successful, returning the resulting Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 = $sSearchString not a String.
;                  @Error 1 @Extended 5 = $atFindFormat not an Array.
;                  @Error 1 @Extended 6 = First element in $atFindFormat not an Object.
;                  @Error 1 @Extended 7 = $oRange not an Object.
;                  @Error 1 @Extended 8 = $oRange has no data selected.
;                  @Error 1 @Extended 9 = $oLastFind not an Object, or failed to retrieve starting position from $oRange.
;                  @Error 1 @Extended 10 = $oLastFind incorrect Object type.
;                  @Error 1 @Extended 11 = $bExhaustive not a Boolean.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When a search is performed inside of a selection, the search may miss any footnotes/ Endnotes/ Frames contained in that selection as the text of these are counted as being located at the very end/beginning of a Document, thus if you are searching in the center of a document, the search will begin in the center, reach the end of the selection, and stop, never reaching the foot/Endnotes etc.
;                  If $bExhaustive is called with True, the search continues until the whole document has been searched, but, if the search has many hits, this could slow the search considerably. There is no use setting this to True in a full document search.
; Related .......: _LOWriter_SearchDescriptorCreate, _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange, _LOWriter_FindFormatModifyAlignment, _LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation, _LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak, _LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace, _LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout, _LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocFindNext(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, $atFindFormat = Null, $oRange = Null, $oLastFind = Null, $bExhaustive = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResult, $oRangeRegion, $oResultRegion, $oText, $oFindRange

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($atFindFormat <> Null) And Not IsArray($atFindFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($atFindFormat <> Null) And (UBound($atFindFormat) > 0) And Not IsObj($atFindFormat[0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If IsArray($atFindFormat) Then $oSrchDescript.setSearchAttributes($atFindFormat)

	If ($oRange <> Null) And Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If ($oRange = Null) Then
		$oRange = $oDoc.getText.createTextCursor()
		$oRange.gotoStart(False)
		$oRange.gotoEnd(True)
	EndIf

	If ($oRange.IsCollapsed()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	If ($oLastFind = Null) Then ; If Last find is not set, then set FindRange to Range beginning or end, depending on SearchBackwards value.
		$oFindRange = ($oSrchDescript.SearchBackwards() = False) ? ($oRange.Start()) : ($oRange.End())

	Else ; If Last find is set, set search start for beginning or end of last result, depending SearchBackwards value.
		If Not IsObj($oLastFind) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not ($oLastFind.supportsService("com.sun.star.text.TextCursor")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		; If Search Backwards is False, then retrieve the end of the last result's range, else get the Start.
		$oFindRange = ($oSrchDescript.SearchBackwards() = False) ? ($oLastFind.End()) : ($oLastFind.Start())
	EndIf

	If Not IsBool($bExhaustive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

	$oSrchDescript.SearchString = $sSearchString

	$oResult = $oDoc.findNext($oFindRange, $oSrchDescript)

	While IsObj($oResult)
		If IsObj($oResult) Then ;  If there is a result, test to see if the result is past the selected area.

			$oRangeRegion = $oRange
			$oResultRegion = $oResult
			$oText = $oDoc.Text

			If $oRange.Text.supportsService("com.sun.star.text.TextFrame") Then
				$oRangeRegion = $oRange.TextFrame.getAnchor() ; If Range is in a TextFrame, convert its position to a range in the document

			ElseIf $oRange.Text.supportsService("com.sun.star.text.Footnote") Or $oRange.Text.supportsService("com.sun.star.text.Endnote") Then
				$oRangeRegion = $oRange.Text.Anchor()
			EndIf

			If $oResult.Text.supportsService("com.sun.star.text.TextFrame") Then
				$oResultRegion = $oResult.TextFrame.getAnchor() ; If Result is in a TextFrame, convert its position to a range in the document

			ElseIf $oResult.Text.supportsService("com.sun.star.text.Footnote") Or $oResult.Text.supportsService("com.sun.star.text.Endnote") Then
				$oResultRegion = $oResult.Text.Anchor()
			EndIf

			If $oRange.Text.supportsService("com.sun.star.text.TextFrame") And $oResult.Text.supportsService("com.sun.star.text.TextFrame") Then
				If ($oDoc.Text.compareRegionEnds($oResultRegion, $oRangeRegion) = 0) Then ; If both Range and Result are in a Text Frame, test if they are in the same one.
					$oResultRegion = $oResult ; If They are, then compare the regions of that text frame.
					$oRangeRegion = $oRange
					$oText = $oRange.Text() ; Must use the corresponding Text Object for that TextFrame as Region Compare can only compare regions contained in the same Text Object region.
				EndIf

			ElseIf $oResult.Text.supportsService("com.sun.star.text.Footnote") Or $oResult.Text.supportsService("com.sun.star.text.Endnote") And _
					$oRange.Text.supportsService("com.sun.star.text.Footnote") Or $oRange.Text.supportsService("com.sun.star.text.Endnote") Then
				If ($oDoc.Text.compareRegionEnds($oResultRegion, $oRangeRegion) = 0) Then ;  If both Range and Result are in a Text Frame, test if they are in the same one.
					$oResultRegion = $oResult ; If They are, then compare the regions of that text frame.
					$oRangeRegion = $oRange
					$oText = $oRange.Text() ; Must use the corresponding Text Object for that Foot/Endnote as Region Compare can only compare regions contained in the same Text Object region.
				EndIf
			EndIf

			If ($oText.compareRegionEnds($oResultRegion, $oRangeRegion) = -1) Then ; If Compare = -1, result is past range.
				If ($bExhaustive = False) Then
					$oResult = Null ; If Result is past the selection set Result to Null, but only if not doing an exhaustive search.
					ExitLoop

				Else ; If $bExhaustive is True, then update the find range.
					$oFindRange = $oResult.End()
				EndIf

			Else ; If Result is within range, exit While loop.
				ExitLoop
			EndIf
		EndIf

		$oResult = $oDoc.findNext($oFindRange, $oSrchDescript)
	WEnd

	Return (IsObj($oResult)) ? (SetError($__LO_STATUS_SUCCESS, 1, $oResult)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocFindNext

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenProp
; Description ...: Set, Retrieve, or reset a Document's General Properties.
; Syntax ........: _LOWriter_DocGenProp(ByRef $oDoc[, $sNewAuthor = Null[, $iRevisions = Null[, $iEditDuration = Null[, $bApplyUserData = Null[, $bResetUserData = False]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sNewAuthor          - [optional] Default is Null. The new author of the document, can be set separately, but must be set to a string if $bResetUserData is called with True.
;                  $iRevisions          - [optional] Default is Null. How often the document was edited and saved.
;                  $iEditDuration       - [optional] Default is Null. The total time of editing the document (in seconds).
;                  $bApplyUserData      - [optional] Default is Null. If True, the user-specific settings saved within a document will be loaded with the document.
;                  $bResetUserData      - [optional] Default is False. If True, clears the document properties, such that it appears the document has just been created. Resets several attributes at once. See remarks.
; Return values .: Success: Integer or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 0 Return 2 = Success. Document Properties were successfully Reset.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters, except $bResetUserData, as it is not a setting.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sNewAuthor not a String and $bResetUserData called with True.
;                  @Error 1 @Extended 3 = $sNewAuthor not a String.
;                  @Error 1 @Extended 4 = $iRevisions not an Integer.
;                  @Error 1 @Extended 5 = $iEditDuration not an Integer.
;                  @Error 1 @Extended 6 = $bApplyUserData not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error retrieving Document Settings Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error retrieving Document Properties Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sNewAuthor
;                  |                               2 = Error setting $iRevisions
;                  |                               4 = Error setting $iEditDuration
;                  |                               8 = Error setting $bApplyUserData
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Setting $bResetUserData to True resets several attributes at once, as follows:
;                  - Author is set to $sNewAuthor parameter, ($sNewAuthor MUST be set to a string).
;                  - CreationDate is set to the current date and time;
;                  - ModifiedBy is cleared, ModificationDate is cleared;
;                  - PrintedBy is cleared; PrintDate is cleared;
;                  - EditingDuration is cleared; EditingCycles is set to 1.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGenProp(ByRef $oDoc, $sNewAuthor = Null, $iRevisions = Null, $iEditDuration = Null, $bApplyUserData = Null, $bResetUserData = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocProp, $oSettings
	Local $iError = 0
	Local $avGenProp[4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSettings = $oDoc.createInstance("com.sun.star.text.DocumentSettings")
	If Not IsObj($oSettings) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bResetUserData = True) Then
		If Not IsString($sNewAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDocProp.resetUserData($sNewAuthor)

		Return SetError($__LO_STATUS_SUCCESS, 0, 2)
	EndIf

	If __LO_VarsAreNull($sNewAuthor, $iRevisions, $iEditDuration, $bApplyUserData) Then
		__LO_ArrayFill($avGenProp, $oDocProp.Author(), $oDocProp.EditingCycles(), $oDocProp.EditingDuration(), $oSettings.getPropertyValue("ApplyUserData"))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGenProp)
	EndIf

	If ($sNewAuthor <> Null) Then
		If Not IsString($sNewAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDocProp.Author = $sNewAuthor
		$iError = ($oDocProp.Author() = $sNewAuthor) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRevisions <> Null) Then
		If Not IsInt($iRevisions) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDocProp.EditingCycles = $iRevisions
		$iError = ($oDocProp.EditingCycles() = $iRevisions) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iEditDuration <> Null) Then
		If Not IsInt($iEditDuration) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oDocProp.EditingDuration = $iEditDuration
		$iError = ($oDocProp.EditingDuration() = $iEditDuration) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bApplyUserData <> Null) Then
		If Not IsBool($bApplyUserData) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oSettings.setPropertyValue("ApplyUserData", $bApplyUserData)
		$iError = ($oSettings.getPropertyValue("ApplyUserData") = $bApplyUserData) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocGenProp

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenPropCreation
; Description ...: Set or Retrieve a Document's General Creation Properties.
; Syntax ........: _LOWriter_DocGenPropCreation(ByRef $oDoc[, $sAuthor = Null[, $tDateStruct = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sAuthor             - [optional] Default is Null. The initial author of the document.
;                  $tDateStruct         - [optional] Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sAuthor not a String.
;                  @Error 1 @Extended 3 = $tDateStruct not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error retrieving Document Properties Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sAuthor
;                  |                               2 = Error setting $tDateStruct
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_DateStructCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGenPropCreation(ByRef $oDoc, $sAuthor = Null, $tDateStruct = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocProp
	Local $iError = 0
	Local $avCreate[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sAuthor, $tDateStruct) Then
		__LO_ArrayFill($avCreate, $oDocProp.Author(), $oDocProp.CreationDate())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avCreate)
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDocProp.Author = $sAuthor
		$iError = ($oDocProp.Author() = $sAuthor) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDocProp.CreationDate = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oDocProp.CreationDate(), $tDateStruct)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocGenPropCreation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenPropModification
; Description ...: Set or Retrieve a Document's General Modification Properties.
; Syntax ........: _LOWriter_DocGenPropModification(ByRef $oDoc[, $sModifiedBy = Null[, $tDateStruct = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sModifiedBy         - [optional] Default is Null. The name of the last user who modified the document.
;                  $tDateStruct         - [optional] Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sModifiedBy not a String.
;                  @Error 1 @Extended 3 = $tDateStruct not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error retrieving Document Properties Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sModifiedBy
;                  |                               2 = Error setting $tDateStruct
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_DateStructCreate, _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGenPropModification(ByRef $oDoc, $sModifiedBy = Null, $tDateStruct = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocProp
	Local $iError = 0
	Local $avMod[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sModifiedBy, $tDateStruct) Then
		__LO_ArrayFill($avMod, $oDocProp.ModifiedBy(), $oDocProp.ModificationDate())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avMod)
	EndIf

	If ($sModifiedBy <> Null) Then
		If Not IsString($sModifiedBy) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDocProp.ModifiedBy = $sModifiedBy
		$iError = ($oDocProp.ModifiedBy() = $sModifiedBy) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDocProp.ModificationDate = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oDocProp.ModificationDate(), $tDateStruct)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocGenPropModification

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenPropPrint
; Description ...: Set or Retrieve a Document's General Printed By Properties.
; Syntax ........: _LOWriter_DocGenPropPrint(ByRef $oDoc[, $sPrintedBy = Null[, $tDateStruct = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sPrintedBy          - [optional] Default is Null. The name of the person who most recently printed the document.
;                  $tDateStruct         - [optional] Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sPrintedBy not a String.
;                  @Error 1 @Extended 3 = $tDateStruct not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error retrieving Document Properties Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sPrintedBy
;                  |                               2 = Error setting $tDateStruct
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_DateStructCreate, _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGenPropPrint(ByRef $oDoc, $sPrintedBy = Null, $tDateStruct = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocProp
	Local $iError = 0
	Local $avPrint[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sPrintedBy, $tDateStruct) Then
		__LO_ArrayFill($avPrint, $oDocProp.PrintedBy(), $oDocProp.PrintDate())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPrint)
	EndIf

	If ($sPrintedBy <> Null) Then
		If Not IsString($sPrintedBy) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDocProp.PrintedBy = $sPrintedBy
		$iError = ($oDocProp.PrintedBy() = $sPrintedBy) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDocProp.PrintDate = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oDocProp.PrintDate(), $tDateStruct)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocGenPropPrint

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenPropTemplate
; Description ...: Set or Retrieve a Document's General Template Properties.
; Syntax ........: _LOWriter_DocGenPropTemplate(ByRef $oDoc[, $sTemplateName = Null[, $sTemplateURL = Null[, $tDateStruct = Null]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sTemplateName       - [optional] Default is Null. The name of the template from which the document was created. The value is an empty string if the document was not created from a template or if it was detached from the template
;                  $sTemplateURL        - [optional] Default is Null. The URL of the template from which the document was created. The value is an empty string if the document was not created from a template or if it was detached from the template.
;                  $tDateStruct         - [optional] Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sTemplateName not a String.
;                  @Error 1 @Extended 3 = $sTemplateURL not a String.
;                  @Error 1 @Extended 4 = $tDateStruct not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error retrieving Document Properties Object.
;                  @Error 3 @Extended 2 = Error converting Computer path to LibreOffice URL.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sTemplateName
;                  |                               2 = Error setting $sTemplateURL
;                  |                               4 = Error setting $tDateStruct
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_DateStructCreate, _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGenPropTemplate(ByRef $oDoc, $sTemplateName = Null, $sTemplateURL = Null, $tDateStruct = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocProp
	Local $iError = 0
	Local $avTemplate[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sTemplateName, $sTemplateURL, $tDateStruct) Then
		__LO_ArrayFill($avTemplate, $oDocProp.TemplateName(), _LO_PathConvert($oDocProp.TemplateURL(), $LO_PATHCONV_PCPATH_RETURN), _
				$oDocProp.TemplateDate())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTemplate)
	EndIf

	If ($sTemplateName <> Null) Then
		If Not IsString($sTemplateName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDocProp.TemplateName = $sTemplateName
		$iError = ($oDocProp.TemplateName() = $sTemplateName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sTemplateURL <> Null) Then
		If Not IsString($sTemplateURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$sTemplateURL = _LO_PathConvert($sTemplateURL, $LO_PATHCONV_OFFICE_RETURN)
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oDocProp.TemplateURL = $sTemplateURL
		$iError = ($oDocProp.TemplateURL() = $sTemplateURL) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDocProp.TemplateDate = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oDocProp.TemplateDate(), $tDateStruct)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocGenPropTemplate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetCounts
; Description ...: Returns the various counts contained in a document, such a paragraph, word etc.
; Syntax ........: _LOWriter_DocGetCounts(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1 dimension array.
;                  @Error 0 @Extended 0 Return Array = Success. A 1 dimension, 0 based, 9 row Array of Integers, in the order described in remarks.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve Document Statistics Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Returns a 1 dimension array with the following counts in this order: Page count; Line Count; Paragraph Count; Word Count; Character Count; NonWhiteSpace Character Count; Table Count; Image Count; Object Count.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGetCounts(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aiCounts[9], $avDocStats

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	__LO_ArrayFill($aiCounts, $oDoc.CurrentController.PageCount(), $oDoc.CurrentController.LineCount(), $oDoc.ParagraphCount(), _
			$oDoc.WordCount(), $oDoc.CharacterCount())

	$avDocStats = $oDoc.DocumentProperties.DocumentStatistics()
	If Not IsArray($avDocStats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($avDocStats) - 1
		If ($avDocStats[$i].Name() = "NonWhitespaceCharacterCount") Then $aiCounts[5] = $avDocStats[$i].Value()
		If ($avDocStats[$i].Name() = "TableCount") Then $aiCounts[6] = $avDocStats[$i].Value()
		If ($avDocStats[$i].Name() = "ImageCount") Then $aiCounts[7] = $avDocStats[$i].Value()
		If ($avDocStats[$i].Name() = "ObjectCount") Then $aiCounts[8] = $avDocStats[$i].Value()
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, $aiCounts)
EndFunc   ;==>_LOWriter_DocGetCounts

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetName
; Description ...: Retrieve the document's name.
; Syntax ........: _LOWriter_DocGetName(ByRef $oDoc[, $bReturnFull = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bReturnFull         - [optional] Default is False. If True, the full window title is returned, such as is used by Autoit window related functions.
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
; Remarks .......: If $bReturnFull is True, the return value will be like: "<Writer Doc name>.<extension> — LibreOffice Writer" e.g. "Testing.odt — LibreOffice Writer".
;                  Else the return value will be like: "<Writer Doc name>.<extension>", e.g. "Testing.odt"
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGetName(ByRef $oDoc, $bReturnFull = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
EndFunc   ;==>_LOWriter_DocGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetPath
; Description ...: Returns a Document's current save path.
; Syntax ........: _LOWriter_DocGetPath(ByRef $oDoc[, $bReturnLibreURL = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bReturnLibreURL     - [optional] Default is False. If True, returns a path in LibreOffice URL format, else False returns a regular Windows path.
; Return values .: Success: String
;                  @Error 0 @Extended 0 Return String = Success. Returning the document's save path as a String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bReturnLibreURL not a Boolean.
;                  @Error 1 @Extended 3 = Document has no save path.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error converting LibreOffice URL to Computer path format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LO_PathConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGetPath(ByRef $oDoc, $bReturnLibreURL = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sPath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnLibreURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.hasLocation() Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$sPath = $oDoc.URL()

	If Not $bReturnLibreURL Then
		$sPath = _LO_PathConvert($sPath, $LO_PATHCONV_PCPATH_RETURN)
		If (@error) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $sPath)
EndFunc   ;==>_LOWriter_DocGetPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHasPath
; Description ...: Returns whether a document has been saved to a location already or not.
; Syntax ........: _LOWriter_DocHasPath(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
Func _LOWriter_DocHasPath(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bHasPath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bHasPath = $oDoc.hasLocation()
	If Not IsBool($bHasPath) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bHasPath)
EndFunc   ;==>_LOWriter_DocHasPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocIsActive
; Description ...: Tests if called document is the active document of other LibreOffice windows.
; Syntax ........: _LOWriter_DocIsActive(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if document is the currently active LibreOffice window. See remarks.
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
Func _LOWriter_DocIsActive(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsActive

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsActive = $oDoc.CurrentController.Frame.isActive()
	If Not IsBool($bIsActive) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsActive)
EndFunc   ;==>_LOWriter_DocIsActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocIsModified
; Description ...: Test whether the document has been modified since being created or since the last save.
; Syntax ........: _LOWriter_DocIsModified(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
Func _LOWriter_DocIsModified(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsMod

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsMod = $oDoc.isModified()
	If Not IsBool($bIsMod) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsMod)
EndFunc   ;==>_LOWriter_DocIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocIsReadOnly
; Description ...: Tests whether a document is opened in ReadOnly mode.
; Syntax ........: _LOWriter_DocIsReadOnly(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
Func _LOWriter_DocIsReadOnly(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsReadOnly

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsReadOnly = $oDoc.isReadOnly()
	If Not IsBool($bIsReadOnly) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsReadOnly)
EndFunc   ;==>_LOWriter_DocIsReadOnly

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocMaximize
; Description ...: Maximize or restore a document.
; Syntax ........: _LOWriter_DocMaximize(ByRef $oDoc[, $bMaximize = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bMaximize           - [optional] Default is Null. If True, document window is maximized, else if False, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean.
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully maximized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMaximize called with Null, returning boolean indicating if Document is currently maximized (True) or not (False).
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bMaximize not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether the Document is Maximized.
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
Func _LOWriter_DocMaximize(ByRef $oDoc, $bMaximize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $bIsMax

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
EndFunc   ;==>_LOWriter_DocMaximize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocMinimize
; Description ...: Minimize or restore a document.
; Syntax ........: _LOWriter_DocMinimize(ByRef $oDoc[, $bMinimize = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bMinimize           - [optional] Default is Null. If True, document window is minimized, else if False, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully minimized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMinimize called with Null, returning boolean indicating if Document is currently minimized (True) or not (False).
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bMinimize not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether the Document is Minimized.
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
Func _LOWriter_DocMinimize(ByRef $oDoc, $bMinimize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
EndFunc   ;==>_LOWriter_DocMinimize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocOpen
; Description ...: Open an existing Writer Document, returning its object identifier.
; Syntax ........: _LOWriter_DocOpen($sFilePath[, $bConnectIfOpen = True[, $bHidden = Null[, $bReadOnly = Null[, $sPassword = Null[, $bLoadAsTemplate = Null[, $sFilterName = Null]]]]]])
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
;                  @Error 1 @Extended 2 = Error converting filepath to URL path.
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
; Related .......: _LOWriter_DocCreate, _LOWriter_DocClose, _LOWriter_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocOpen($sFilePath, $bConnectIfOpen = True, $bHidden = Null, $bReadOnly = Null, $sPassword = Null, $bLoadAsTemplate = Null, $sFilterName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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

	If $bConnectIfOpen Then $oDoc = _LOWriter_DocConnect($LO_DOC_CONNECT_MODE_SEARCH_PATH, $sFilePath, True)
	If IsObj($oDoc) Then Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))

	$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOWriter_DocOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPosAndSize
; Description ...: Reposition and resize a document window.
; Syntax ........: _LOWriter_DocPosAndSize(ByRef $oDoc[, $iX = Null[, $iY = Null[, $iWidth = Null[, $iHeight = Null]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
Func _LOWriter_DocPosAndSize(ByRef $oDoc, $iX = Null, $iY = Null, $iWidth = Null, $iHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
EndFunc   ;==>_LOWriter_DocPosAndSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrint
; Description ...: Print a document using the specified settings.
; Syntax ........: _LOWriter_DocPrint(ByRef $oDoc[, $iCopies = 1[, $bCollate = True[, $vPages = "ALL"[, $bWait = True[, $iDuplexMode = $LOW_PRINT_DUPLEX_OFF[, $sPrinter = ""[, $sFilePathName = ""]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iCopies             - [optional] Default is 1. Specifies the number of copies to print.
;                  $bCollate            - [optional] Default is True. Advises the printer to collate the pages of the copies.
;                  $vPages              - [optional] Default is "ALL". Specifies which pages to print. See remarks.
;                  $bWait               - [optional] Default is True. If True, the corresponding print request will be executed synchronous. Default is to use synchronous print mode.
;                  $iDuplexMode         - [optional] (0-3) Default is $__g_iDuplexOFF. Determines the duplex mode for the print job. See Constants, $LOW_PRINT_DUPLEX_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPrinter            - [optional] Default is "". Printer name. If left blank, or if printer name is not found, default printer is used.
;                  $sFilePathName       - [optional] Default is "". Specifies the name of a file to print to. Creates a .prn file at the given Path. Must include the desired path destination with file name.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success Document was successfully printed.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $iCopies not a Integer.
;                  @Error 1 @Extended 3 = $bCollate not a Boolean.
;                  @Error 1 @Extended 4 = $vPages not an Integer or String.
;                  @Error 1 @Extended 5 = $vPages contains invalid characters, a-z, or a period(.).
;                  @Error 1 @Extended 6 = $bWait not a Boolean.
;                  @Error 1 @Extended 7 = $iDuplexMode not an Integer, less than 0 or greater than 3. See Constants, $LOW_PRINT_DUPLEX_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 = $sPrinter not a String.
;                  @Error 1 @Extended 9 = $sFilePathName not a
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error creating "Printer Name" property.
;                  @Error 2 @Extended 2 = Error creating "Copies" property.
;                  @Error 2 @Extended 3 = Error creating "Collate" property.
;                  @Error 2 @Extended 4 = Error creating "Wait" property.
;                  @Error 2 @Extended 5 = Error creating "DuplexMode" property.
;                  @Error 2 @Extended 6 = Error creating "Pages" property.
;                  @Error 2 @Extended 7 = Error creating "PrintToFile" property.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error converting PrintToFile Path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Based on OOoCalc UDF Print function by GMK.
;                  $vPages range can be called as entered in the user interface, as follows: "1-4,10" to print the pages 1 to 4 and 10. Default is "ALL". Must be in String format to accept more than just a single page number. e.g. This will work: "1-6,12,27" This will not 1-6,12,27. This will work: "7", This will also: 7.
;                  Setting $bWait to True is highly recommended. Otherwise following actions (as e.g. closing the Document) can fail.
; Related .......: _LO_PrintersGetNamesAlt, _LO_PrintersGetNames, _LOWriter_DocPrintSizeSettings, _LOWriter_DocPrintPageSettings, _LOWriter_DocPrintMiscSettings, _LOWriter_DocPrintIncludedSettings
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocPrint(ByRef $oDoc, $iCopies = 1, $bCollate = True, $vPages = "ALL", $bWait = True, $iDuplexMode = $LOW_PRINT_DUPLEX_OFF, $sPrinter = "", $sFilePathName = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
	If Not __LO_IntIsBetween($iDuplexMode, $LOW_PRINT_DUPLEX_OFF, $LOW_PRINT_DUPLEX_SHORT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
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
EndFunc   ;==>_LOWriter_DocPrint

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrintIncludedSettings
; Description ...: Set or Retrieve setting related to what is included in printing.
; Syntax ........: _LOWriter_DocPrintIncludedSettings(ByRef $oDoc[, $bGraphics = Null[, $bControls = Null[, $bDrawings = Null[, $bTables = Null[, $bHiddenText = Null]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bGraphics           - [optional] Default is Null. If True, the graphics contained in the document are printed.
;                  $bControls           - [optional] Default is Null. If True, the form control fields contained in the document are printed.
;                  $bDrawings           - [optional] Default is Null. If True, the drawings contained in the document are printed.
;                  $bTables             - [optional] Default is Null. If True, the Tables contained in the document are printed.
;                  $bHiddenText         - [optional] Default is Null. If True, prints text that is marked as hidden.
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bGraphics not a Boolean.
;                  @Error 1 @Extended 3 = $bControls not a Boolean.
;                  @Error 1 @Extended 4 = $bDrawings not a Boolean.
;                  @Error 1 @Extended 5 = $bTables not a Boolean.
;                  @Error 1 @Extended 6 = $bHiddenText not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Failed to create "com.sun.star.text.DocumentSettings" Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bGraphics
;                  |                               2 = Error setting $bControls
;                  |                               4 = Error setting $bDrawings
;                  |                               8 = Error setting $bTables
;                  |                               16 = Error setting $bHiddenText
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_DocPrintSizeSettings, _LOWriter_DocPrintPageSettings, _LOWriter_DocPrintMiscSettings
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocPrintIncludedSettings(ByRef $oDoc, $bGraphics = Null, $bControls = Null, $bDrawings = Null, $bTables = Null, $bHiddenText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSettings
	Local $iError = 0
	Local $abPrintSettings[5]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSettings = $oDoc.createInstance("com.sun.star.text.DocumentSettings")
	If Not IsObj($oSettings) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If __LO_VarsAreNull($bGraphics, $bControls, $bDrawings, $bTables, $bHiddenText) Then
		__LO_ArrayFill($abPrintSettings, $oSettings.getPropertyValue("PrintGraphics"), $oSettings.getPropertyValue("PrintControls"), _
				$oSettings.getPropertyValue("PrintDrawings"), $oSettings.getPropertyValue("PrintTables"), $oSettings.getPropertyValue("PrintHiddenText"))

		Return SetError($__LO_STATUS_SUCCESS, 1, $abPrintSettings)
	EndIf

	If ($bGraphics <> Null) Then
		If Not IsBool($bGraphics) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oSettings.setPropertyValue("PrintGraphics", $bGraphics)
		$iError = ($oSettings.getPropertyValue("PrintGraphics") = $bGraphics) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bControls <> Null) Then
		If Not IsBool($bControls) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oSettings.setPropertyValue("PrintControls", $bControls)
		$iError = ($oSettings.getPropertyValue("PrintControls") = $bControls) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bDrawings <> Null) Then
		If Not IsBool($bDrawings) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oSettings.setPropertyValue("PrintDrawings", $bDrawings)
		$iError = ($oSettings.getPropertyValue("PrintDrawings") = $bDrawings) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bTables <> Null) Then
		If Not IsBool($bTables) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oSettings.setPropertyValue("PrintTables", $bTables)
		$iError = ($oSettings.getPropertyValue("PrintTables") = $bTables) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bHiddenText <> Null) Then
		If Not IsBool($bHiddenText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oSettings.setPropertyValue("PrintHiddenText", $bHiddenText)
		$iError = ($oSettings.getPropertyValue("PrintHiddenText") = $bHiddenText) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocPrintIncludedSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrintMiscSettings
; Description ...: Set or Retrieve Miscellaneous Printing related settings.
; Syntax ........: _LOWriter_DocPrintMiscSettings(ByRef $oDoc[, $iPaperOrient = Null[, $sPrinterName = Null[, $iCommentsMode = Null[, $bBrochure = Null[, $bBrochureRTL = Null[, $bReversed = Null]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iPaperOrient        - [optional] (0-1) Default is Null. The orientation of the paper. See Constants, $LOW_PAPER_ORIENT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPrinterName        - [optional] Default is Null. The name of the printer to send print jobs to.
;                  $iCommentsMode       - [optional] (0-3) Default is Null. If and where to print comments in the document. See Constants, $LOW_PRINT_NOTES_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bBrochure           - [optional] Default is Null. If True, prints the document in brochure format.
;                  $bBrochureRTL        - [optional] Default is Null. If True, prints the document in brochure Right to Left format.
;                  $bReversed           - [optional] Default is Null. If True, prints pages in reverse order.
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $iPaperOrient not an Integer, less than 0 or greater than 1. See Constants, $LOW_PAPER_ORIENT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 3 = $sPrinterName not a string.
;                  @Error 1 @Extended 4 = $iCommentsMode not an Integer, less than 0 or greater than 3. See Constants, $LOW_PRINT_NOTES_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 = $bBrochure not a Boolean.
;                  @Error 1 @Extended 6 = $bBrochureRTL not a Boolean.
;                  @Error 1 @Extended 7 = $bReversed not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Failed to create "com.sun.star.text.DocumentSettings" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error retrieving setting value of "CanSetPaperOrientation" from Printer.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPaperOrient
;                  |                               2 = Error setting $sPrinterName
;                  |                               4 = Error setting $iCommentsMode
;                  |                               8 = Error setting $bBrochure
;                  |                               16 = Error setting $bBrochureRTL
;                  |                               32 = Error setting $bReversed
;                  --Printer Related Errors--
;                  @Error 5 @Extended 1 = Printer does not allow changing paper orientation.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_DocPrintSizeSettings, _LOWriter_DocPrintPageSettings, _LOWriter_DocPrintIncludedSettings
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocPrintMiscSettings(ByRef $oDoc, $iPaperOrient = Null, $sPrinterName = Null, $iCommentsMode = Null, $bBrochure = Null, $bBrochureRTL = Null, $bReversed = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__STR_STRIPLEADING = 1, $__STR_STRIPTRAILING = 2
	Local $iError = 0
	Local $oSettings
	Local $bCanSetPaperOrientation = False
	Local $aoSetting[1]
	Local $avPrintSettings[6]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSettings = $oDoc.createInstance("com.sun.star.text.DocumentSettings")
	If Not IsObj($oSettings) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If __LO_VarsAreNull($iPaperOrient, $sPrinterName, $iCommentsMode, $bBrochure, $bBrochureRTL, $bReversed) Then
		__LO_ArrayFill($avPrintSettings, __LOWriter_GetPrinterSetting($oDoc, "PaperOrientation"), _
				__LOWriter_GetPrinterSetting($oDoc, "Name"), $oSettings.getPropertyValue("PrintAnnotationMode"), _
				$oSettings.getPropertyValue("PrintProspect"), $oSettings.getPropertyValue("PrintProspectRTL"), _
				$oSettings.getPropertyValue("PrintReversed"))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPrintSettings)
	EndIf

	$bCanSetPaperOrientation = __LOWriter_GetPrinterSetting($oDoc, "CanSetPaperOrientation")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iPaperOrient <> Null) Then
		If Not __LO_IntIsBetween($iPaperOrient, $LOW_PAPER_ORIENT_PORTRAIT, $LOW_PAPER_ORIENT_LANDSCAPE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If $bCanSetPaperOrientation Then
			$aoSetting[0] = __LO_SetPropertyValue("PaperOrientation", $iPaperOrient)
			$oDoc.setPrinter($aoSetting)
			$iError = (__LOWriter_GetPrinterSetting($oDoc, "PaperOrientation") = $iPaperOrient) ? ($iError) : (BitOR($iError, 1))

		Else

			Return SetError($__LO_STATUS_PRINTER_RELATED_ERROR, 1, 0)
		EndIf
	EndIf

	If ($sPrinterName <> Null) Then
		If Not IsString($sPrinterName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$sPrinterName = StringStripWS(StringStripWS($sPrinterName, $__STR_STRIPTRAILING), $__STR_STRIPLEADING)
		$aoSetting[0] = __LO_SetPropertyValue("Name", $sPrinterName)
		$oDoc.setPrinter($aoSetting)
		$iError = (__LOWriter_GetPrinterSetting($oDoc, "Name") = $sPrinterName) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iCommentsMode <> Null) Then
		If Not __LO_IntIsBetween($iCommentsMode, $LOW_PRINT_NOTES_NONE, $LOW_PRINT_NOTES_NEXT_PAGE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oSettings.setPropertyValue("PrintAnnotationMode", $iCommentsMode)
		$iError = ($oSettings.getPropertyValue("PrintAnnotationMode") = $iCommentsMode) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bBrochure <> Null) Then
		If Not IsBool($bBrochure) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oSettings.setPropertyValue("PrintProspect", $bBrochure)
		$iError = ($oSettings.getPropertyValue("PrintProspect") = $bBrochure) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bBrochureRTL <> Null) Then
		If Not IsBool($bBrochureRTL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oSettings.setPropertyValue("PrintProspectRTL", $bBrochureRTL)
		$iError = ($oSettings.getPropertyValue("PrintProspectRTL") = $bBrochureRTL) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bReversed <> Null) Then
		If Not IsBool($bReversed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oSettings.setPropertyValue("PrintReversed", $bReversed)
		$iError = ($oSettings.getPropertyValue("PrintReversed") = $bReversed) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocPrintMiscSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrintPageSettings
; Description ...: Set or Retrieve settings Page related print settings.
; Syntax ........: _LOWriter_DocPrintPageSettings(ByRef $oDoc[, $bBlackOnly = Null[, $bLeftOnly = Null[, $bRightOnly = Null[, $bBackground = Null[, $bEmptyPages = Null]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bBlackOnly          - [optional] Default is Null. If True, prints all text in black only.
;                  $bLeftOnly           - [optional] Default is Null. If True, prints only Left(Even) pages. See remarks.
;                  $bRightOnly          - [optional] Default is Null. If True, prints only Right(Odd) pages. See remarks.
;                  $bBackground         - [optional] Default is Null. If True, prints colors and objects that are inserted to the background of the page.
;                  $bEmptyPages         - [optional] Default is Null. If True, automatically inserted blank pages are printed.
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bBlackOnly not a Boolean.
;                  @Error 1 @Extended 3 = $bLeftOnly not a Boolean.
;                  @Error 1 @Extended 4 = $bRightOnly not a Boolean.
;                  @Error 1 @Extended 5 = $bBackground not a Boolean.
;                  @Error 1 @Extended 6 = $bEmptyPages not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Failed to create "com.sun.star.text.DocumentSettings" Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bBlackOnly
;                  |                               2 = Error setting $bLeftOnly
;                  |                               4 = Error setting $bRightOnly
;                  |                               8 = Error setting $bBackground
;                  |                               16 = Error setting $bEmptyPages
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If both $bLeftOnly and $bRightOnly are True, both Left and Right pages are printed.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_DocPrintSizeSettings, _LOWriter_DocPrintMiscSettings, _LOWriter_DocPrintIncludedSettings
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocPrintPageSettings(ByRef $oDoc, $bBlackOnly = Null, $bLeftOnly = Null, $bRightOnly = Null, $bBackground = Null, $bEmptyPages = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oSettings
	Local $abPrintSettings[5]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oSettings = $oDoc.createInstance("com.sun.star.text.DocumentSettings")
	If Not IsObj($oSettings) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If __LO_VarsAreNull($bBlackOnly, $bLeftOnly, $bRightOnly, $bBackground, $bEmptyPages) Then
		__LO_ArrayFill($abPrintSettings, $oSettings.getPropertyValue("PrintBlackFonts"), $oSettings.getPropertyValue("PrintLeftPages"), _
				$oSettings.getPropertyValue("PrintRightPages"), $oSettings.getPropertyValue("PrintPageBackground"), $oSettings.getPropertyValue("PrintEmptyPages"))

		Return SetError($__LO_STATUS_SUCCESS, 1, $abPrintSettings)
	EndIf

	If ($bBlackOnly <> Null) Then
		If Not IsBool($bBlackOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oSettings.setPropertyValue("PrintBlackFonts", $bBlackOnly)
		$iError = ($oSettings.getPropertyValue("PrintBlackFonts") = $bBlackOnly) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bLeftOnly <> Null) Then
		If Not IsBool($bLeftOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oSettings.setPropertyValue("PrintLeftPages", $bLeftOnly)
		$iError = ($oSettings.getPropertyValue("PrintLeftPages") = $bLeftOnly) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bRightOnly <> Null) Then
		If Not IsBool($bRightOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oSettings.setPropertyValue("PrintRightPages", $bRightOnly)
		$iError = ($oSettings.getPropertyValue("PrintRightPages") = $bRightOnly) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bBackground <> Null) Then
		If Not IsBool($bBackground) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oSettings.setPropertyValue("PrintPageBackground", $bBackground)
		$iError = ($oSettings.getPropertyValue("PrintPageBackground") = $bBackground) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bEmptyPages <> Null) Then
		If Not IsBool($bEmptyPages) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oSettings.setPropertyValue("PrintEmptyPages", $bEmptyPages)
		$iError = ($oSettings.getPropertyValue("PrintEmptyPages") = $bEmptyPages) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocPrintPageSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrintSizeSettings
; Description ...: Set or Retrieve Print Paper size settings.
; Syntax ........: _LOWriter_DocPrintSizeSettings(ByRef $oDoc[, $iPaperFormat = Null[, $iPaperWidth = Null[, $iPaperHeight = Null]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iPaperFormat        - [optional] (0-8) Default is Null. Specifies a predefined paper size or if the paper size is a user-defined size. See constants, $LOW_PAPER_FORMAT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iPaperWidth         - [optional] Default is Null. Specifies the size of the paper in Hundredths of a Millimeter (HMM). Can be a custom value or one of the constants, $LOW_PAPER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
;                  $iPaperHeight        - [optional] Default is Null. Specifies the size of the paper in Hundredths of a Millimeter (HMM). Can be a custom value or one of the constants, $LOW_PAPER_HEIGHT_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
; Return values .: Success: 1 or Array.
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $iPaperFormat not an Integer, less than 0 or greater than 8. See constants, $LOW_PAPER_FORMAT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 3 = $iPaperWidth not an Integer.
;                  @Error 1 @Extended 4 = $iPaperHeight not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Failed to create "com.sun.star.awt.Size" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve Printer setting "CanSetPaperFormat".
;                  @Error 3 @Extended 2 = Failed to retrieve Printer setting "CanSetPaperSize".
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPaperFormat
;                  |                               2 = Error setting $iPaperWidth
;                  |                               4 = Error setting $iPaperHeight
;                  --Printer Related Errors--
;                  @Error 5 @Extended 1 = Printer doesn't allow paper format to be set.
;                  @Error 5 @Extended 2 = Printer doesn't allow paper size to be set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Due to slight inaccuracies in unit conversion, there may be False errors thrown while attempting to set paper size.
;                  For some reason, setting $iPaperWidth and $iPaperHeight modifies the document page size also.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_UnitConvert, _LOWriter_DocPrintPageSettings, _LOWriter_DocPrintMiscSettings, _LOWriter_DocPrintIncludedSettings
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocPrintSizeSettings(ByRef $oDoc, $iPaperFormat = Null, $iPaperWidth = Null, $iPaperHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bCanSetPaperFormat = False, $bCanSetPaperSize = False
	Local $iError = 0
	Local $tSize
	Local $aoSetting[1]
	Local $aiPrintSettings[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iPaperFormat, $iPaperWidth, $iPaperHeight) Then
		__LO_ArrayFill($aiPrintSettings, __LOWriter_GetPrinterSetting($oDoc, "PaperFormat"), _
				_LO_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Width(), $LO_CONVERT_UNIT_TWIPS_HMM), _
				_LO_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Height(), $LO_CONVERT_UNIT_TWIPS_HMM))

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiPrintSettings)
	EndIf

	$bCanSetPaperFormat = __LOWriter_GetPrinterSetting($oDoc, "CanSetPaperFormat")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$bCanSetPaperSize = __LOWriter_GetPrinterSetting($oDoc, "CanSetPaperSize")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If ($iPaperFormat <> Null) Then
		If Not __LO_IntIsBetween($iPaperFormat, $LOW_PAPER_FORMAT_A3, $LOW_PAPER_FORMAT_USER_DEFINED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		If $bCanSetPaperFormat Then
			$aoSetting[0] = __LO_SetPropertyValue("PaperFormat", $iPaperFormat)
			$oDoc.setPrinter($aoSetting)
			$iError = (__LOWriter_GetPrinterSetting($oDoc, "PaperFormat") = $iPaperFormat) ? ($iError) : (BitOR($iError, 1))

		Else

			Return SetError($__LO_STATUS_PRINTER_RELATED_ERROR, 1, 0)
		EndIf
	EndIf

	If ($iPaperWidth <> Null) Or ($iPaperHeight <> Null) Then
		If $bCanSetPaperSize Then
			If Not IsInt($iPaperWidth) And ($iPaperWidth <> Null) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
			If Not IsInt($iPaperHeight) And ($iPaperHeight <> Null) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			; Set in Hundredths of a Millimeter (HMM) but retrieved in TWIPS
			$tSize = __LO_CreateStruct("com.sun.star.awt.Size")
			If Not IsObj($tSize) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tSize.Width = ($iPaperWidth = Null) ? (_LO_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Width(), $LO_CONVERT_UNIT_TWIPS_HMM)) : ($iPaperWidth)
			$tSize.Height = ($iPaperWidth = Null) ? (_LO_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Height(), $LO_CONVERT_UNIT_TWIPS_HMM)) : ($iPaperHeight)
			$aoSetting[0] = __LO_SetPropertyValue("PaperSize", $tSize)
			$oDoc.setPrinter($aoSetting)

			$iError = (__LO_VarsAreNull($iPaperWidth)) ? ($iError) : (__LO_IntIsBetween(_LO_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Width(), $LO_CONVERT_UNIT_TWIPS_HMM), $iPaperWidth - 2, $iPaperWidth + 2)) ? ($iError) : (BitOR($iError, 2))
			$iError = (__LO_VarsAreNull($iPaperHeight)) ? ($iError) : (__LO_IntIsBetween(_LO_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Height(), $LO_CONVERT_UNIT_TWIPS_HMM), $iPaperHeight - 2, $iPaperHeight + 2)) ? ($iError) : (BitOR($iError, 4))

		Else

			Return SetError($__LO_STATUS_PRINTER_RELATED_ERROR, 2, 0)
		EndIf
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocPrintSizeSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocRedo
; Description ...: Perform one Redo action for a document.
; Syntax ........: _LOWriter_DocRedo(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Related .......: _LOWriter_DocUndo, _LOWriter_DocRedoIsPossible, _LOWriter_DocRedoGetAllActionTitles, _LOWriter_DocRedoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocRedo(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isRedoPossible()) Then
		$oDoc.UndoManager.Redo()

		Return SetError($__LO_STATUS_SUCCESS, 1, 0)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOWriter_DocRedo

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocRedoClear
; Description ...: Clear all Redo Actions in the Redo Action List.
; Syntax ........: _LOWriter_DocRedoClear(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared all Redo Actions.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This will silently fail if there are any _LOWriter_DocUndoActionBegin still active.
; Related .......: _LOWriter_DocUndoClear, _LOWriter_DocUndoReset
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocRedoClear(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.clearRedo()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocRedoClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocRedoCurActionTitle
; Description ...: Retrieve the current Redo action Title.
; Syntax ........: _LOWriter_DocRedoCurActionTitle(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Related .......: _LOWriter_DocRedo, _LOWriter_DocRedoGetAllActionTitles
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocRedoCurActionTitle(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sRedoAction

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sRedoAction = $oDoc.UndoManager.getCurrentRedoActionTitle()
	If ($sRedoAction = Null) Then $sRedoAction = ""
	If Not IsString($sRedoAction) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sRedoAction)
EndFunc   ;==>_LOWriter_DocRedoCurActionTitle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocRedoGetAllActionTitles
; Description ...: Retrieve all available Redo action Titles.
; Syntax ........: _LOWriter_DocRedoGetAllActionTitles(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Related .......: _LOWriter_DocRedo, _LOWriter_DocRedoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocRedoGetAllActionTitles(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asTitles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asTitles = $oDoc.UndoManager.getAllRedoActionTitles()
	If Not IsArray($asTitles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asTitles), $asTitles)
EndFunc   ;==>_LOWriter_DocRedoGetAllActionTitles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocRedoIsPossible
; Description ...: Test whether a Redo action is available to perform for a document.
; Syntax ........: _LOWriter_DocRedoIsPossible(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;                  @Error 0 @Extended 0 Return Boolean = If the document has a redo action to perform, True is returned, else False.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query if a Redo is possible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocRedo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocRedoIsPossible(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsRedoPoss

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsRedoPoss = $oDoc.UndoManager.isRedoPossible()
	If Not IsBool($bIsRedoPoss) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $bIsRedoPoss)
EndFunc   ;==>_LOWriter_DocRedoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocReplaceAll
; Description ...: Replace all instances of a search.
; Syntax ........: _LOWriter_DocReplaceAll(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, $sReplaceString[, $atFindFormat = Null[, $atReplaceFormat = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - A String of text or a Regular Expression to Search for.
;                  $sReplaceString      - A String of text or a Regular Expression to replace any results with.
;                  $atFindFormat        - [optional] Default is Null. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
;                  $atReplaceFormat     - [optional] Default is Null. An Array of Formatting property values to replace any results with.
; Return values .: Success: Integer
;                  @Error 0 @Extended 0 Return Integer = Success. Search and Replace was successful, returning number of replacements made.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 = $sSearchString not a String.
;                  @Error 1 @Extended 5 = $sReplaceString not a String.
;                  @Error 1 @Extended 6 = $atFindFormat not an Array.
;                  @Error 1 @Extended 7 = $atReplaceFormat not an Array.
;                  @Error 1 @Extended 8 = First Element of $atFindFormat not an Object.
;                  @Error 1 @Extended 9 = First Element of $atReplaceFormat not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: In order for $atReplaceFormat to be applied to replacements, $bSearchPropValues must be True in the Search descriptor. I'm not sure why.
;                  Calling $bBackwards with True can cause issues with Find and Replace using formats, perhaps other things as well.
; Related .......: _LOWriter_SearchDescriptorCreate, _LOWriter_DocFindAll, _LOWriter_DocFindNext, _LOWriter_DocFindAllInRange, _LOWriter_DocReplaceAllInRange, _LOWriter_FindFormatModifyAlignment, _LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation, _LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak, _LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace, _LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout, _LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocReplaceAll(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, $sReplaceString, $atFindFormat = Null, $atReplaceFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iReplacements

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sReplaceString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($atFindFormat <> Null) And Not IsArray($atFindFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($atReplaceFormat <> Null) And Not IsArray($atReplaceFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If ($atFindFormat <> Null) And (UBound($atFindFormat) > 0) And Not IsObj($atFindFormat[0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If ($atReplaceFormat <> Null) And (UBound($atReplaceFormat) > 0) And Not IsObj($atReplaceFormat[0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	If IsArray($atFindFormat) Then $oSrchDescript.setSearchAttributes($atFindFormat)
	If IsArray($atReplaceFormat) Then $oSrchDescript.setReplaceAttributes($atReplaceFormat)

	$oSrchDescript.SearchString = $sSearchString
	$oSrchDescript.ReplaceString = $sReplaceString

	$iReplacements = $oDoc.replaceAll($oSrchDescript)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iReplacements)
EndFunc   ;==>_LOWriter_DocReplaceAll

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocReplaceAllInRange
; Description ...: Replace all instances of a search within a selection.
; Syntax ........: _LOWriter_DocReplaceAllInRange(ByRef $oDoc, ByRef $oSrchDescript, ByRef $oRange, $sSearchString, $sReplaceString[, $atFindFormat = Null[, $atReplaceFormat = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $oRange              - A Range, such as a cursor with Data selected, to perform the search within.
;                  $sSearchString       - A String of text or a regular expression to search for.
;                  $sReplaceString      - A String of text or a regular expression to replace any results with.
;                  $atFindFormat        - [optional] Default is Null. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
;                  $atReplaceFormat     - [optional] Default is Null. An Array of Formatting property values to replace any results with.
; Return values .: Success: Integer
;                  @Error 0 @Extended 0 Return Integer = Success. Search and Replace was successful, returning number of replacements.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 = $oRange not an Object.
;                  @Error 1 @Extended 5 = $oRange contains no selected Data.
;                  @Error 1 @Extended 6 = $sSearchString not a String.
;                  @Error 1 @Extended 7 = $sReplaceString not a String.
;                  @Error 1 @Extended 8 = $atFindFormat not an Array.
;                  @Error 1 @Extended 9 = $atReplaceFormat not an Array.
;                  @Error 1 @Extended 10 = First Element in $atFindFormat not a Property Object.
;                  @Error 1 @Extended 11 = First Element in $atReplaceFormat not a Property Object.
;                  @Error 1 @Extended 12 = Paragraph Style Name called in $sReplaceString does not exist.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error creating backup of ViewCursor location and selection.
;                  @Error 2 @Extended 2 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 3 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error Finding all results in Range.
;                  @Error 3 @Extended 2 = Error searching for property values.
;                  @Error 3 @Extended 3 = Error finding temporary property to use.
;                  @Error 3 @Extended 4 = Error retrieving current selection and ViewCursor position.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibreOffice does not offer a method to replace only results within a selection, consequently I have had to create my own. This function sometimes uses the "FindAllInRange" function, so any errors with Find/Replace formatting causing deletions will cause problems here. As best as I can tell all options for find and replace should be available, Formatting, Paragraph styles etc.
;                  If formatting is not being search or applied, I use a dispatch command to Find and Replace. However if formatting is being searched or added, A second method is used, which begins with the "FindAllInRange" function to find all matching results, then temporarily applies a normally unused property to the applicable results (CharFlash or CharShadingValue), and then add that temporary property to the Formatting array to search for, then Replace all results. And finally removing the temporary property value again.
;                  Replacing Paragraph Styles doesn't work with a dispatch command, so I use the "FindAllInRange" function, and then manually apply the new Paragraph Style.
;                  In order for $atReplaceFormat to be applied to replacements, $bSearchPropValues must be True in the Search descriptor. I'm not sure why.
;                  Calling $bBackwards with True can cause issues with Find and Replace using formats, perhaps other things as well.
; Related .......: _LOWriter_SearchDescriptorCreate, _LOWriter_DocFindAll, _LOWriter_DocFindNext, _LOWriter_DocFindAllInRange, _LOWriter_DocReplaceAll, _LOWriter_FindFormatModifyAlignment, _LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation, _LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak, _LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace, _LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout, _LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocReplaceAllInRange(ByRef $oDoc, ByRef $oSrchDescript, ByRef $oRange, $sSearchString, $sReplaceString, $atFindFormat = Null, $atReplaceFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__LOW_ALG1_ABSOLUTE = 0, $__LOW_ALG1_REGEXP = 1, $__LOW_ALG1_APPROXIMATE = 2 ; com.sun.star.util.SearchAlgorithms
	Local Const $__LOW_ALG2_ABSOLUTE = 1, $__LOW_ALG2_REGEXP = 2, $__LOW_ALG2_APPROXIMATE = 3 ; com.sun.star.util.SearchAlgorithms2
	Local Const $__LOW_SRCH_COMMAND_REPLACE_ALL = 3 ; See srchitem.hxx and https://thebiasplanet.blogspot.com/2022/06/writerunosearchoff.html
	Local Const $__LOW_TRANSLIT_FLAG_NONE = 0, $__LOW_TRANSLIT_FLAG_IGNORE_CASE = 256 ; com.sun.star.i18n:TransliterationModules
	Local Const $__LOW_SEARCHFLAG_NORM_WORD_ONLY = 16, $__LOW_SEARCHFLAG_SELECTION = 2048, $__LOW_SEARCHFLAG_LEV_RELAXED = 65536 ; See com,sun,star,util,SearchFlags, srchitem.hxx, https://thebiasplanet.blogspot.com/2022/06/writerunosearchoff.html
	Local $aoResults[0]
	Local $atArgs[12], $atFormats[1], $atOrigFormats[1]
	Local $oServiceManager, $oDispatcher, $oTempSrchDescript, $oResults, $oSelection
	Local $iResults, $iSrchFlags = $__LOW_SEARCHFLAG_SELECTION, $iTranslitFlags = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($oRange.IsCollapsed()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsString($sSearchString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsString($sReplaceString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If ($atFindFormat <> Null) And Not IsArray($atFindFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If ($atReplaceFormat <> Null) And Not IsArray($atReplaceFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If ($atFindFormat <> Null) And (UBound($atFindFormat) > 0) And Not IsObj($atFindFormat[0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
	If ($atReplaceFormat <> Null) And Not (UBound($atReplaceFormat) > 0) And Not IsObj($atReplaceFormat[0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
	If ($oSrchDescript.SearchStyles() = True) And Not _LOWriter_ParStyleExists($oDoc, $sReplaceString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

	$aoResults = _LOWriter_DocFindAllInRange($oDoc, $oSrchDescript, $sSearchString, $oRange, $atFindFormat)
	$iResults = @extended
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Error performing search

	If IsArray($atFindFormat) Or IsArray($atReplaceFormat) Then ; Search or replace using formats, use my temporary properties method.
		$oTempSrchDescript = $oDoc.createSearchDescriptor()
		If Not IsObj($oTempSrchDescript) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		With $oTempSrchDescript
			.SearchBackwards = False
			.SearchCaseSensitive = False
			.SearchWords = False
			.SearchRegularExpression = True
			.SearchStyles = True
			.ValueSearch = False
		EndWith

		; Use these as Temp values. CharFlash, or CharShadingValue.
		$atFormats[0] = __LO_SetPropertyValue("CharFlash", True)
		$atOrigFormats[0] = __LO_SetPropertyValue("CharFlash", False)

		$oTempSrchDescript.setSearchAttributes($atFormats)
		$oTempSrchDescript.SearchString = ".*"

		$oResults = $oDoc.findAll($oTempSrchDescript)
		If Not IsObj($oResults) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		If ($oResults.getCount() > 0) Then ; If CharFlash is present, try to find another unused property.
			$atFormats[0] = __LO_SetPropertyValue("CharShadingValue", 28)
			$atOrigFormats[0] = __LO_SetPropertyValue("CharShadingValue", 0)
			$oTempSrchDescript.setSearchAttributes($atFormats)
			$oResults = $oDoc.findAll($oTempSrchDescript)
			If Not IsObj($oResults) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		EndIf

		If ($oResults.getCount() > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oDoc.getUndoManager.enterUndoContext("Find and Replace")

		For $i = 0 To $iResults - 1
			$aoResults[$i].setPropertyValue($atFormats[0].Name(), $atFormats[0].Value())

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		If IsArray($atFindFormat) Then
			__LOWriter_FindFormatAddSetting($atFindFormat, $atFormats[0])

		Else
			$atFindFormat = $atFormats
		EndIf

		$oSrchDescript.setSearchAttributes($atFindFormat)
		If IsArray($atReplaceFormat) Then $oSrchDescript.setReplaceAttributes($atReplaceFormat)

		$oSrchDescript.SearchString = $sSearchString
		$oSrchDescript.ReplaceString = $sReplaceString

		$oDoc.replaceAll($oSrchDescript)

		$oTempSrchDescript.setReplaceAttributes($atOrigFormats)
		$oTempSrchDescript.ReplaceString = "&"
		$oTempSrchDescript.ValueSearch = True

		$oDoc.replaceAll($oTempSrchDescript)

		$oDoc.getUndoManager.leaveUndoContext()

		Return SetError($__LO_STATUS_SUCCESS, 0, $iResults)

	ElseIf ($oSrchDescript.SearchStyles() = True) Then ; Paragraph Style replacement (Dispatch doesn't work for these).
		$oDoc.getUndoManager.enterUndoContext("Replace Style " & $sSearchString)

		For $i = 0 To $iResults - 1
			$aoResults[$i].ParaStyleName = $sReplaceString

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
		$oDoc.getUndoManager.leaveUndoContext()

		Return SetError($__LO_STATUS_SUCCESS, 1, $iResults)

	Else ; Use Dispatch.
		; Backup the ViewCursor location and selection.
		$oSelection = $oDoc.getCurrentSelection()
		If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		; Move the View Cursor to the input range and select it.
		$oDoc.CurrentController.Select($oRange)

		$oServiceManager = $oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
		If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$iSrchFlags = (($oSrchDescript.SearchSimilarity() = True) And ($oSrchDescript.SearchSimilarityRelax() = True)) ? (BitOR($iSrchFlags, $__LOW_SEARCHFLAG_LEV_RELAXED)) : ($iSrchFlags)
		$iSrchFlags = (($oSrchDescript.SearchWords() = True)) ? (BitOR($iSrchFlags, $__LOW_SEARCHFLAG_NORM_WORD_ONLY)) : ($iSrchFlags)

		$iTranslitFlags = ($oSrchDescript.SearchCaseSensitive() = True) ? ($__LOW_TRANSLIT_FLAG_NONE) : ($__LOW_TRANSLIT_FLAG_IGNORE_CASE)

		$atArgs[0] = __LO_SetPropertyValue("SearchItem.AlgorithmType", (($oSrchDescript.SearchSimilarity() = True) ? ($__LOW_ALG1_APPROXIMATE) : (($oSrchDescript.SearchRegularExpression() = True) ? ($__LOW_ALG1_REGEXP) : ($__LOW_ALG1_ABSOLUTE))))
		$atArgs[1] = __LO_SetPropertyValue("SearchItem.AlgorithmType2", (($oSrchDescript.SearchSimilarity() = True) ? ($__LOW_ALG2_APPROXIMATE) : (($oSrchDescript.SearchRegularExpression() = True) ? ($__LOW_ALG2_REGEXP) : ($__LOW_ALG2_ABSOLUTE))))
		$atArgs[2] = __LO_SetPropertyValue("SearchItem.Backward", $oSrchDescript.SearchBackwards())
		$atArgs[3] = __LO_SetPropertyValue("SearchItem.ChangedChars", $oSrchDescript.SearchSimilarityExchange())
		$atArgs[4] = __LO_SetPropertyValue("SearchItem.Command", $__LOW_SRCH_COMMAND_REPLACE_ALL)
		$atArgs[5] = __LO_SetPropertyValue("SearchItem.DeletedChars", $oSrchDescript.SearchSimilarityRemove())
		$atArgs[6] = __LO_SetPropertyValue("SearchItem.InsertedChars", $oSrchDescript.SearchSimilarityAdd())
		$atArgs[7] = __LO_SetPropertyValue("SearchItem.Pattern", $oSrchDescript.SearchStyles())
		$atArgs[8] = __LO_SetPropertyValue("SearchItem.ReplaceString", $sReplaceString)
		$atArgs[9] = __LO_SetPropertyValue("SearchItem.SearchFlags", $iSrchFlags)
		$atArgs[10] = __LO_SetPropertyValue("SearchItem.SearchString", $sSearchString)
		$atArgs[11] = __LO_SetPropertyValue("SearchItem.TransliterateFlags", $iTranslitFlags)

		$oDispatcher.executeDispatch($oDoc.CurrentController, ".uno:ExecuteSearch", "", 0, $atArgs)

		; Restore the ViewCursor to its previous location.
		$oDoc.CurrentController.Select($oSelection)

		Return SetError($__LO_STATUS_SUCCESS, 2, $iResults)
	EndIf
EndFunc   ;==>_LOWriter_DocReplaceAllInRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocSave
; Description ...: Save any changes made to a Document.
; Syntax ........: _LOWriter_DocSave(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Document Successfully saved.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Document is ReadOnly or Document has no save location, try SaveAs.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocExport, _LOWriter_DocSaveAs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocSave(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDoc.hasLocation Or $oDoc.isReadOnly Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oDoc.store()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocSaveAs
; Description ...: Save a Document with the specified file name to the path specified with any parameters called.
; Syntax ........: _LOWriter_DocSaveAs(ByRef $oDoc, $sFilePath[, $sFilterName = ""[, $bOverwrite = Null[, $sPassword = Null]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFilePath           - Full path to save the document to, including Filename and extension.
;                  $sFilterName         - [optional] Default is "". The filter name. Calling "" (blank string), means the filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .odt extension is used instead, with the filter name of "writer8".
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
;                  @Error 2 @Extended 1 = Error creating FilterName Property
;                  @Error 2 @Extended 2 = Error creating Overwrite Property
;                  @Error 2 @Extended 3 = Error creating Password Property
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error Converting Path to/from L.O. URL
;                  @Error 3 @Extended 2 = Error retrieving FilterName.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Alters original save path (if there was one) to the new path.
; Related .......: _LOWriter_DocExport, _LOWriter_DocSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocSaveAs(ByRef $oDoc, $sFilePath, $sFilterName = "", $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aProperties[1]
	Local $sSavePath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFilterName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$sFilePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFilterName = "") Or ($sFilterName = " ") Then $sFilterName = __LOWriter_FilterNameGet($sFilePath)
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
EndFunc   ;==>_LOWriter_DocSaveAs

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocSelection
; Description ...: Set or Retrieve the current Document selection(s).
; Syntax ........: _LOWriter_DocSelection(ByRef $oDoc[, $oObj = Null[, $bReturnMultiAsObj = False]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObj                - Default is Null. A selectable object. A Text Cursor with text selected, A ViewCursor with Text Selected, a Table Cursor with cells selected, a Shape or Frame Object, etc.
;                  $bReturnMultiAsObj   - [optional] Default is False. If True, when Multiple selections are present, they will be returned as a single Object. See Remarks.
; Return values .: Success: 1, Object or Array
;                  @Error 0 @Extended -6 Return Object = Success. The current selection is within a Table, returning a Table Cursor Object.
;                  @Error 0 @Extended -5 Return Object = Success. The current selection was a Frame, returning a Frame Object.
;                  @Error 0 @Extended -4 Return Object = Success. The current selection was a Shape, returning a Shape Object.
;                  @Error 0 @Extended -3 Return Object = Success. The current selection was a Chart or other OLE object, returning the Object.
;                  @Error 0 @Extended -2 Return Object = Success. The current selection was an Image, returning a Image Object.
;                  @Error 0 @Extended -1 Return Object = Success. The current selection is multiple disconnected selections and $bReturnMultiAsObj was True, Returning a single Object.
;                  @Error 0 @Extended 0 Return 1 = Success. Object called in $oObj successfully selected.
;                  @Error 0 @Extended 1 Return Object = Success. The current selection is a single span of text. Returning a Text Cursor.
;                  @Error 0 @Extended ? Return Array = Success. The current selection is multiple disconnected selections. Returning an Array of Text Cursors. @Extended is set to number of results.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $oObj not an Object.
;                  @Error 1 @Extended 3 = $bReturnMultiAsObj not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Failed to create a TextCursor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve current selection.
;                  @Error 3 @Extended 2 =There is no text selected.
;                  @Error 3 @Extended 3 = Failed to retrieve count of multiple selections.
;                  @Error 3 @Extended 4 = Failed to identify current selection.
;                  @Error 3 @Extended 5 = Failed to select object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call $oObj with Null to retrieve the current selection.
;                  If there are multiple selections present, the default behaviour of this function is to create a TextCursor for each selection and return them in an Array. When $bReturnMultiAsObj is True, the single multi-selection Object (com.sun.star.text.TextRanges) is returned, this is useless to the user (unless they use the API commands themselves to retrieve the individual selections), but can be used to restore the previous selections by calling the returned Object in this function.
;                  Presently, I have no way to set multiple selections at a time other than the above mentioned method.
;                  When multiple selections are present, one returned cursor will usually be the present position of the ViewCursor in the current selection.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocSelection(ByRef $oDoc, $oObj = Null, $bReturnMultiAsObj = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bSelect
	Local $oSelection, $oCursor
	Local $aoSelections[0]
	Local $iCount

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($bReturnMultiAsObj <> Null) And Not IsBool($bReturnMultiAsObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($oObj) Then
		$oSelection = $oDoc.getCurrentSelection()
		If Not IsObj($oSelection) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If $oSelection.supportsService("com.sun.star.text.TextTableCursor") Then

			Return SetError($__LO_STATUS_SUCCESS, -6, $oSelection)

		ElseIf $oSelection.supportsService("com.sun.star.text.TextFrame") Then

			Return SetError($__LO_STATUS_SUCCESS, -5, $oSelection)

		ElseIf $oSelection.supportsService("com.sun.star.drawing.Shapes") Then

			Return SetError($__LO_STATUS_SUCCESS, -4, $oSelection)

		ElseIf $oSelection.supportsService("com.sun.star.text.TextEmbeddedObject") Then ; Chart? Function etc?

			Return SetError($__LO_STATUS_SUCCESS, -3, $oSelection)

		ElseIf $oSelection.supportsService("com.sun.star.text.TextGraphicObject") Then ; Image? etc?

			Return SetError($__LO_STATUS_SUCCESS, -2, $oSelection)

		ElseIf $oSelection.supportsService("com.sun.star.text.TextRanges") Then
			If ($oSelection.Count() > 1) Then
				If $bReturnMultiAsObj Then

					Return SetError($__LO_STATUS_SUCCESS, -1, $oSelection)

				Else
					$iCount = $oSelection.getCount()
					If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

					ReDim $aoSelections[$iCount]

					For $i = 0 To $iCount - 1
						$aoSelections[$i] = $oSelection.getByIndex($i).Text.createTextCursorByRange($oSelection.getByIndex($i))
						If Not IsObj($aoSelections[$i]) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
					Next

					Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoSelections)
				EndIf

			Else
				$oCursor = $oSelection.getByIndex(0).Text.createTextCursorByRange($oSelection.getByIndex(0))
				If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
				If $oCursor.isCollapsed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				Return SetError($__LO_STATUS_SUCCESS, 1, $oCursor)
			EndIf

		Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
		EndIf
	EndIf

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$bSelect = $oDoc.CurrentController.Select($oObj)

	If ($oObj.supportsService("com.sun.star.text.TextTable")) Then
		$oCursor = _LOWriter_CursorViewCursorGetObj($oDoc)
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		_LOWriter_CursorMove($oCursor, $LOW_VIEWCUR_GOTO_END, 1, True) ; Move and select to End of cell
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		_LOWriter_CursorMove($oCursor, $LOW_VIEWCUR_GOTO_END, 1, True) ; Move and select to End of Table
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndIf

	If ($bSelect = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocSelection

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocToFront
; Description ...: Bring the called document to the front of the other windows.
; Syntax ........: _LOWriter_DocToFront(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
Func _LOWriter_DocToFront(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.toFront()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocToFront

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndo
; Description ...: Perform one Undo action for a document.
; Syntax ........: _LOWriter_DocUndo(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Related .......: _LOWriter_DocUndoIsPossible, _LOWriter_DocUndoGetAllActionTitles, _LOWriter_DocUndoCurActionTitle, _LOWriter_DocRedo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocUndo(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isUndoPossible()) Then
		$oDoc.UndoManager.Undo()

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOWriter_DocUndo

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoActionBegin
; Description ...: Begin an Undo Action group.
; Syntax ........: _LOWriter_DocUndoActionBegin(ByRef $oDoc[, $sName = "AU3LO-Automation"])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
;                  _LOWriter_DocUndoActionEnd must be called after this function before this undo group will become available in the Undo Action list.
;                  _LOWriter_DocUndoActionBegin can be nested, e.g. call this function multiple times without ending the first undo action, but only the last group that is ended with _LOWriter_DocUndoActionEnd will appear.
; Related .......: _LOWriter_DocUndoActionEnd
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocUndoActionBegin(ByRef $oDoc, $sName = "AU3LO-Automation")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.UndoManager.enterUndoContext($sName)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocUndoActionBegin

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoActionEnd
; Description ...: End the last started Undo Action Group.
; Syntax ........: _LOWriter_DocUndoActionEnd(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully ended the last Undo Action group recording.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This stops the grouping of actions into the last created Undo Action Group.
; Related .......: _LOWriter_DocUndoActionBegin
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocUndoActionEnd(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.leaveUndoContext()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocUndoActionEnd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoClear
; Description ...: Clear all Undo and Redo Actions in the Undo/Redo Action List.
; Syntax ........: _LOWriter_DocUndoClear(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared all Undo and Redo Actions.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This will silently fail if there are any _LOWriter_DocUndoActionBegin still active.
; Related .......: _LOWriter_DocRedoClear
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocUndoClear(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.clear()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocUndoClear

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoCurActionTitle
; Description ...: Retrieve the current Undo action Title.
; Syntax ........: _LOWriter_DocUndoCurActionTitle(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Related .......: _LOWriter_DocUndo, _LOWriter_DocUndoGetAllActionTitles
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocUndoCurActionTitle(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sUndoAction

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sUndoAction = $oDoc.UndoManager.getCurrentUndoActionTitle()
	If ($sUndoAction = Null) Then $sUndoAction = ""
	If Not IsString($sUndoAction) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sUndoAction)
EndFunc   ;==>_LOWriter_DocUndoCurActionTitle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoGetAllActionTitles
; Description ...: Retrieve all available Undo action Titles.
; Syntax ........: _LOWriter_DocUndoGetAllActionTitles(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Related .......: _LOWriter_DocUndo, _LOWriter_DocUndoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocUndoGetAllActionTitles(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asTitles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asTitles = $oDoc.UndoManager.getAllUndoActionTitles()
	If Not IsArray($asTitles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asTitles), $asTitles)
EndFunc   ;==>_LOWriter_DocUndoGetAllActionTitles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoIsPossible
; Description ...: Test whether a Undo action is available to perform for a document.
; Syntax ........: _LOWriter_DocUndoIsPossible(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;                  @Error 0 @Extended 0 Return Boolean = If the document has an undo action to perform, True is returned, else False.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query if an Undo is possible.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocUndo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocUndoIsPossible(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsUndoPoss

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsUndoPoss = $oDoc.UndoManager.isUndoPossible()
	If Not IsBool($bIsUndoPoss) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $bIsUndoPoss)
EndFunc   ;==>_LOWriter_DocUndoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoReset
; Description ...: Reset the UndoManager.
; Syntax ........: _LOWriter_DocUndoReset(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully reset the undo manager.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling this function does the following: remove all locks from the undo manager; closes all open undo group actions, clears all undo actions, clears all redo actions.
; Related .......: _LOWriter_DocUndoClear, _LOWriter_DocRedoClear
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocUndoReset(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc.UndoManager.reset()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocUndoReset

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocVisible
; Description ...: Set or retrieve the current visibility of a document.
; Syntax ........: _LOWriter_DocVisible(ByRef $oDoc[, $bVisible = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
Func _LOWriter_DocVisible(ByRef $oDoc, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
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
EndFunc   ;==>_LOWriter_DocVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocZoom
; Description ...: Modify the zoom value for a document.
; Syntax ........: _LOWriter_DocZoom(ByRef $oDoc[, $iZoom = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iZoom               - [optional] (20-600) Default is Null. The zoom percentage.
; Return values .: Success: Integer.
;                  @Error 0 @Extended 0 Return 1 = $iZoom set successfully.
;                  @Error 0 @Extended 1 Return Integer =  All optional parameters were called with Null, returning current zoom value.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $iZoom not an Integer, less than 20 or greater than 600.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve current Zoom value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iZoom
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocZoom(ByRef $oDoc, $iZoom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iCurZoom
	Local $oServiceManager, $oDispatcher
	Local $aArgs[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iZoom) Then
		$iCurZoom = $oDoc.CurrentController.ViewSettings.ZoomValue()
		If Not IsInt($iCurZoom) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurZoom)
	EndIf

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iZoom, 20, 600) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$aArgs[0] = __LO_SetPropertyValue("Zoom.Value", $iZoom)
	$aArgs[1] = __LO_SetPropertyValue("Zoom.ValueSet", 28703)
	$aArgs[2] = __LO_SetPropertyValue("Zoom.Type", 0)

	$oDispatcher.executeDispatch($oDoc.CurrentController, ".uno:Zoom", "", 0, $aArgs)
	$iError = ($oDoc.CurrentController.ViewSettings.ZoomValue() = $iZoom) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocZoom
