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
; _LOWriter_DocBookmarkDelete
; _LOWriter_DocBookmarkExists
; _LOWriter_DocBookmarkGetAnchor
; _LOWriter_DocBookmarkGetObj
; _LOWriter_DocBookmarkInsert
; _LOWriter_DocBookmarkModify
; _LOWriter_DocBookmarksGetNames
; _LOWriter_DocClose
; _LOWriter_DocConnect
; _LOWriter_DocConvertTableToText
; _LOWriter_DocConvertTextToTable
; _LOWriter_DocCreate
; _LOWriter_DocCreateTextCursor
; _LOWriter_DocDescription
; _LOWriter_DocExecuteDispatch
; _LOWriter_DocExport
; _LOWriter_DocFindAll
; _LOWriter_DocFindAllInRange
; _LOWriter_DocFindNext
; _LOWriter_DocFooterGetTextCursor
; _LOWriter_DocFormSettings
; _LOWriter_DocGenProp
; _LOWriter_DocGenPropCreation
; _LOWriter_DocGenPropModification
; _LOWriter_DocGenPropPrint
; _LOWriter_DocGenPropTemplate
; _LOWriter_DocGetCounts
; _LOWriter_DocGetName
; _LOWriter_DocGetPath
; _LOWriter_DocGetString
; _LOWriter_DocGetViewCursor
; _LOWriter_DocHasPath
; _LOWriter_DocHeaderGetTextCursor
; _LOWriter_DocHyperlinkInsert
; _LOWriter_DocInsertControlChar
; _LOWriter_DocInsertString
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
; _LOWriter_DocViewCursorGetPosition
; _LOWriter_DocVisible
; _LOWriter_DocZoom
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkDelete
; Description ...: Delete a Bookmark.
; Syntax ........: _LOWriter_DocBookmarkDelete(ByRef $oDoc, ByRef $oBookmark)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oBookmark           - [in/out] an object. A Bookmark Object from a previous _LOWriter_DocBookmarkInsert, or _LOWriter_DocBookmarkGetObj function to delete.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oBookmark not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Attempted to delete Bookmark, but document still contains a Bookmark by that name.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested Bookmark.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocBookmarkInsert, _LOWriter_DocBookmarkGetObj, _LOWriter_DocBookmarksGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkDelete(ByRef $oDoc, ByRef $oBookmark)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sBookmarkName

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oBookmark) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sBookmarkName = $oBookmark.Name()

	$oBookmark.dispose()

	Return (_LOWriter_DocBookmarkExists($oDoc, $sBookmarkName)) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocBookmarkDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkExists
; Description ...: Check if a document contains a Bookmark by name.
; Syntax ........: _LOWriter_DocBookmarkExists(ByRef $oDoc, $sBookmarkName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sBookmarkName       - a string value. The Bookmark name to search for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sBookmarkName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Bookmarks Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the document contains a Bookmark by the called name, then True is returned, Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkExists(ByRef $oDoc, $sBookmarkName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBookmarks

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sBookmarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oBookmarks = $oDoc.getBookmarks()
	If Not IsObj($oBookmarks) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oBookmarks.hasByName($sBookmarkName))
EndFunc   ;==>_LOWriter_DocBookmarkExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkGetAnchor
; Description ...: Retrieve a Bookmark's Anchor cursor Object.
; Syntax ........: _LOWriter_DocBookmarkGetAnchor(ByRef $oBookmark)
; Parameters ....: $oBookmark           - [in/out] an object. A Bookmark Object from a previous _LOWriter_DocBookmarkInsert, or _LOWriter_DocBookmarkGetObj function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oBookmark not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to retrieve Bookmark anchor Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Bookmark Anchor Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Anchor cursor returned is just a Text Cursor placed at the anchor's position.
; Related .......: _LOWriter_DocBookmarkGetObj, _LOWriter_DocBookmarkInsert, _LOWriter_CursorMove, _LOWriter_DocGetString, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkGetAnchor(ByRef $oBookmark)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBookAnchor

	If Not IsObj($oBookmark) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oBookAnchor = $oBookmark.Anchor.Text.createTextCursorByRange($oBookmark.Anchor())
	If Not IsObj($oBookAnchor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oBookAnchor)
EndFunc   ;==>_LOWriter_DocBookmarkGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkGetObj
; Description ...: Retrieve a Bookmark Object by name.
; Syntax ........: _LOWriter_DocBookmarkGetObj(ByRef $oDoc, $sBookmarkName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sBookmarkName       - a string value. The Bookmark name to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sBookmarkName not a String.
;                  @Error 1 @Extended 3 Return 0 = Document does not contain a Bookmark named in $sBookmarkName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve requested Bookmark Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully retrieved requested Bookmark Object. Returning requested Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocBookmarksGetNames, _LOWriter_DocBookmarkModify, _LOWriter_DocBookmarkDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkGetObj(ByRef $oDoc, $sBookmarkName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBookmark

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sBookmarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_DocBookmarkExists($oDoc, $sBookmarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oBookmark = $oDoc.Bookmarks.getByName($sBookmarkName)
	If Not IsObj($oBookmark) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oBookmark)
EndFunc   ;==>_LOWriter_DocBookmarkGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkInsert
; Description ...: Insert a Bookmark into a document.
; Syntax ........: _LOWriter_DocBookmarkInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $sBookmarkName = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $sBookmarkName       - [optional] a string value. Default is Null. The Name of the Bookmark to create. See Remarks.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, which is not supported.
;                  @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $sBookmarkName not a String.
;                  @Error 1 @Extended 6 Return 0 = $sBookmarkName contains illegal characters, /\@:*?";,.# .
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.Bookmark" Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Bookmark into the document. Returning the Bookmark Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the cursor used to insert a Bookmark has text selected, the Bookmark will envelope the text, else the Bookmark will be inserted at a single point.
;                  A Bookmark name cannot contain the following characters: / \ @ : * ? " ; , . #
;                  If the document already contains a Bookmark by the same name, Libre Office adds a digit after the name, such as Bookmark 1, Bookmark 2 etc.
; Related .......: _LOWriter_DocBookmarkModify, _LOWriter_DocBookmarkDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sBookmarkName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBookmark

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oBookmark = $oDoc.createInstance("com.sun.star.text.Bookmark")
	If Not IsObj($oBookmark) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sBookmarkName <> Null) Then
		If Not IsString($sBookmarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If StringRegExp($sBookmarkName, '[/\@:*?";,.#]') Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; Invalid Characters in Name.

		$oBookmark.Name = $sBookmarkName

	Else
		$oBookmark.Name = "Bookmark "
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oBookmark, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oBookmark)
EndFunc   ;==>_LOWriter_DocBookmarkInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkModify
; Description ...: Set or Retrieve a Bookmark's settings.
; Syntax ........: _LOWriter_DocBookmarkModify(ByRef $oBookmark[, $sBookmarkName = Null])
; Parameters ....: $oBookmark           - [in/out] an object. A Bookmark Object from a previous _LOWriter_DocBookmarkInsert, or _LOWriter_DocBookmarkGetObj function.
;                  $sBookmarkName       - [optional] a string value. Default is Null. The new name to rename the bookmark called in $oBookmark.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oBookmark not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sBookmarkName not a String.
;                  @Error 1 @Extended 3 Return 0 = $sBookmarkName contains illegal characters, /\@:*?";,.# .
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sBookmarkName
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Bookmark name successfully modified.
;                  @Error 0 @Extended 0 Return String = Success. All optional parameters were called with Null, returning current Bookmark name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  A Bookmark name cannot contain the following characters: / \ @ : * ? " ; , . #
;                  If the document already contains a Bookmark by the same name, Libre Office adds a digit after the name, such as Bookmark 1, Bookmark 2 etc.
; Related .......: _LOWriter_DocBookmarkGetObj, _LOWriter_DocBookmarkInsert, _LOWriter_DocBookmarkDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkModify(ByRef $oBookmark, $sBookmarkName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oBookmark) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sBookmarkName) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oBookmark.Name())

	If Not IsString($sBookmarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If StringRegExp($sBookmarkName, '[/\@:*?";,.#]') Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Invalid Characters in Name.

	$oBookmark.Name = $sBookmarkName
	$iError = ($oBookmark.Name() = $sBookmarkName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocBookmarkModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarksGetNames
; Description ...: Retrieve an Array of Bookmark names.
; Syntax ........: _LOWriter_DocBookmarksGetNames(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Array of Bookmark Names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Successfully searched for Bookmarks, returning Array of Bookmark names, @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocBookmarkGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarksGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asBookmarkNames[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asBookmarkNames = $oDoc.Bookmarks.getElementNames()
	If Not IsArray($asBookmarkNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asBookmarkNames), $asBookmarkNames)
EndFunc   ;==>_LOWriter_DocBookmarksGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocClose
; Description ...: Close an existing Writer Document, returning its save path if applicable.
; Syntax ........: _LOWriter_DocClose(ByRef $oDoc[, $bSaveChanges = True[, $sSaveName = ""[, $bDeliverOwnership = True]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bSaveChanges        - [optional] a boolean value. Default is True. If True, saves changes if any were made before closing. See remarks.
;                  $sSaveName           - [optional] a string value. Default is "". The file name to save the file as, if the file hasn't been saved before. See Remarks.
;                  $bDeliverOwnership   - [optional] a boolean value. Default is True. If True, deliver ownership of the document Object from the script to LibreOffice, recommended is True.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bSaveChanges not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sSaveName not a String.
;                  @Error 1 @Extended 4 Return 0 = $bDeliverOwnership not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error while creating Filter Name properties.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Path Conversion to L.O. URL Failed.
;                  @Error 3 @Extended 2 Return 0 = Error while retrieving FilterName.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Success, Document was successfully closed, and was saved to the returned file Path.
;                  @Error 0 @Extended 2 Return String = Success, Document was successfully closed, document's changes were saved to its existing location.
;                  @Error 0 @Extended 3 Return String = Success, Document was successfully closed, document either had no changes to save, or $bSaveChanges was called with False. If document had a save location, or if document was saved to a location, it is returned, else an empty string is returned.
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
EndFunc   ;==>_LOWriter_DocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocConnect
; Description ...: Connect to an already opened instance of LibreOffice Writer.
; Syntax ........: _LOWriter_DocConnect($sFile[, $bConnectCurrent = False[, $bConnectAll = False]])
; Parameters ....: $sFile               - a string value. A Full or partial file path, or a full or partial file name. See remarks. Can be an empty string if $bConnectAll or $bConnectCurrent is True.
;                  $bConnectCurrent     - [optional] a boolean value. Default is False. If True, returns the currently active, or last active Document, unless it is not a Text Document.
;                  $bConnectAll         - [optional] a boolean value. Default is False. If True, returns an array containing all open LibreOffice Writer Text Documents. See remarks.
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
;                  @Error 3 @Extended 2 Return 0 = Current Component not a Text Document.
;                  @Error 3 @Extended 3 Return 0 = Error converting path to Libre Office URL.
;                  @Error 3 @Extended 4 Return 0 = No matches found.
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Success, The Object for the current, or last active document is returned.
;                  @Error 0 @Extended ? Return Array = Success, An Array of all open LibreOffice Writer Text documents is returned. See remarks. @Extended is set to number of results.
;                  @Error 0 @Extended 3 Return Object = Success, The Object for the document with matching URL is returned.
;                  @Error 0 @Extended 4 Return Object = Success, The Object for the document with matching Title is returned.
;                  @Error 0 @Extended 5 Return Object = Success, A partial Title or Path search found only one match, returning the Object for the found document.
;                  @Error 0 @Extended ? Return Array = Success, An Array of all matching Libre Text documents from a partial Title or Path search. See remarks. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFile can be either the full Path (Name and extension included; e.g: C:\file\Test.odt Or file:///C:/file/Test.odt) of the document, or the full Title with extension, (e.g: Test.odt), or a partial file path (e.g: file1\file2\Test Or file1\file2 Or file1/file2/ etc.), or a partial name (e.g: test, would match Test1.odt, Test2.docx etc.).
;                  Partial file path searches and file name searches, as well as the connect all option, return arrays with three columns per result. ($aArray[0][3]). each result is stored in a separate row;
;                  -Row 1, Column 0 contains the Object for that document. e.g. $aArray[0][0] = $oDoc
;                  -Row 1, Column 1 contains the Document's full title and extension. e.g. $aArray[0][1] = This Test File.docx
;                  -Row 1, Column 2 contains the document's full file path. e.g. $aArray[0][2] = C:\Folder1\Folder2\This Test File.docx
;                  -Row 2, Column 0 contains the Object for the next document. And so on. e.g. $aArray[1][0] = $oDoc2
; Related .......: _LOWriter_DocOpen, _LOWriter_DocClose, _LOWriter_DocCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocConnect($sFile, $bConnectCurrent = False, $bConnectAll = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local Const $__STR_STRIPLEADING = 1
	Local $aoConnectAll[1], $aoPartNameSearch[1]
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop
	Local $sServiceName = "com.sun.star.text.TextDocument"

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

		Return ($oDoc.supportsService($sServiceName) And Not IsObj($oDoc.Parent())) ? (SetError($__LO_STATUS_SUCCESS, 1, $oDoc)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0))
	EndIf

	If $bConnectAll Then
		ReDim $aoConnectAll[1][3]
		$iCount = 0
		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService($sServiceName) And Not IsObj($oDoc.Parent()) Then ; If Parent is an Object, then Writer doc is a DataBase Form

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

		$iCount = 0 ; partial name or partial URL search
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
EndFunc   ;==>_LOWriter_DocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocConvertTableToText
; Description ...: Convert a Table to Text, separated by a delimiter.
; Syntax ........: _LOWriter_DocConvertTableToText(ByRef $oDoc, ByRef $oTable[, $sDelimiter = @TAB])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $sDelimiter          - [optional] a string value. Default is @TAB. A character to separate each column by, such as a Tab character, etc.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sDelimiter not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve array of CellNames.
;                  @Error 3 @Extended 2 Return 0 = Failed to create a backup of the ViewCursor's current location.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Table was successfully converted to text.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions. See Remarks.
;                  $sDelimiter          - [optional] a string value. Default is @TAB. A character to the text into each column by, such as a Tab etc.
;                  $bHeader             - [optional] a boolean value. Default is False. If True, Formats the first row of the new table as a heading.
;                  $iRepeatHeaderLines  - [optional] an integer value. Default is 0. If greater than 0, then Repeats the first n rows as a header.
;                  $bBorder             - [optional] a boolean value. Default is False. If True, Adds a border to the table and the table cells.
;                  $bDontSplitTable     - [optional] a boolean value. Default is False. If True, Does not divide the table across pages.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sDelimiter not a String.
;                  @Error 1 @Extended 4 Return 0 = $bHeader not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iRepeatHeaderLines not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $bBorder not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bDontSplitTable not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $oCursor is a Table Cursor and is not supported.
;                  @Error 1 @Extended 9 Return 0 = $oCursor has no data selected.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve TextTables Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve array of Table names.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify $oCursor's cursor type.
;                  @Error 3 @Extended 4 Return 0 = Failed to backup ViewCursor's position.
;                  @Error 3 @Extended 5 Return 0 = Failed to retrieve new Table's Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Text was successfully converted to a Table, returning the new Table's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function temporarily moves the ViewCursor to and selects the Text, and then attempts to restore the ViewCursor to its former position.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_ParObjCreateList, _LOWriter_DocConvertTableToText
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

	Return (IsObj($oTable)) ? (SetError($__LO_STATUS_SUCCESS, 0, $oTable)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0))
EndFunc   ;==>_LOWriter_DocConvertTextToTable

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocCreate
; Description ...: Open a new Libre Office Writer Document or Connect to an existing blank, unsaved, writable document.
; Syntax ........: _LOWriter_DocCreate([$bForceNew = True[, $bHidden = False]])
; Parameters ....: $bForceNew           - [optional] a boolean value. Default is True. If True, force opening a new Writer Document instead of checking for a usable blank.
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
; Name ..........: _LOWriter_DocCreateTextCursor
; Description ...: Create a TextCursor Object for future Textcursor related functional use.
; Syntax ........: _LOWriter_DocCreateTextCursor(ByRef $oDoc[, $bCreateAtEnd = True[, $bCreateAtViewCursor = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
;                  $bCreateAtEnd        - [optional] a boolean value. Default is True. If True, creates the new cursor at the end of the Document. Else cursor is created at the beginning.
;                  $bCreateAtViewCursor - [optional] a boolean value. Default is False. If True, create the Text cursor at the document's View Cursor. See Remarks
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bCreateAtEnd not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bCreateAtViewCursor not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bCreateAtEnd and $bCreateAtViewCursor both called with True, set either one to False.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Text Cursor Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve ViewCursor Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to Retrieve Text Object.
;                  @Error 3 @Extended 3 Return 0 = Current ViewCursor is in unknown data type or failed detecting what data type.
;                  --Success--
;                  @Error 0 @Extended ? Return Object = Success, Cursor object was returned. @Extended can be on of the constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3 indicating the current created cursor is in that type of data.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The cursor Created by this function in a text document, is used for inserting text, reading text, etc.
;                  If you set $bCreateAtEnd to False, the new cursor is created at the beginning of the document, True creates the cursor at the very end of the document.
;                  Setting $bCreateAtViewCursor to True will create a Textcursor at the current ViewCursor position.
;                  There are two types of cursors in Word documents. The one you see, called the "ViewCursor", and one you do not see, called a "TextCursor". A "ViewCursor" is the blinking cursor you see when you are editing a Word document, there is only one per document. A "TextCursor" on the other hand, is an invisible cursor used for inserting text etc., into a Writer document. You can have multiple "TextCursors" per document.
; Related .......: _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocCreateTextCursor(ByRef $oDoc, $bCreateAtEnd = True, $bCreateAtViewCursor = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCursor, $oText, $oViewCursor
	Local $iCursorType = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bCreateAtEnd) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bCreateAtViewCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($bCreateAtEnd = True) And ($bCreateAtViewCursor = True) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If ($bCreateAtViewCursor = True) Then
		$oViewCursor = $oDoc.CurrentController.getViewCursor()
		If Not IsObj($oViewCursor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$oText = __LOWriter_CursorGetText($oDoc, $oViewCursor)
		$iCursorType = @extended
		If @error Or Not IsObj($oText) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		If __LO_IntIsBetween($iCursorType, $LOW_CURDATA_BODY_TEXT, $LOW_CURDATA_HEADER_FOOTER) Then
			$oCursor = $oText.createTextCursorByRange($oViewCursor)

		Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; ViewCursor in unknown data type.
		EndIf

	Else
		$oText = $oDoc.getText
		If Not IsObj($oText) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oCursor = $oText.createTextCursor()
		If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$iCursorType = $LOW_CURDATA_BODY_TEXT

		If ($bCreateAtEnd = True) Then
			$oCursor.gotoEnd(False)

		Else
			$oCursor.gotoStart(False)
		EndIf
	EndIf

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, $iCursorType, $oCursor)
EndFunc   ;==>_LOWriter_DocCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocDescription
; Description ...: Set or Retrieve Document Description properties.
; Syntax ........: _LOWriter_DocDescription(ByRef $oDoc[, $sTitle = Null[, $sSubject = Null[, $asKeywords = Null[, $sComments = Null[, $asContributor = Null[, $sCoverage = Null[, $sIdentifier = Null[, $asPublisher = Null[, $asRelation = Null[, $sRights = Null[, $sSource = Null[, $sType = Null]]]]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
;                  $sTitle              - [optional] a string value. Default is Null. The Document's "Title" Property. See Remarks.
;                  $sSubject            - [optional] a string value. Default is Null. The Document's "Subject" Property.
;                  $asKeywords          - [optional] an array of strings. Default is Null. The Document's "Keywords" Property. Input must be a single dimension Array, which will overwrite any keywords previously set. Accepts numbers also. See Remarks.
;                  $sComments           - [optional] a string value. Default is Null. The Document's "Comments" Property.
;                  $asContributor       - [optional] an array of strings. Default is Null. The Document's "Contributor" Property. Input must be a single dimension Array, which will overwrite any values previously set. See Remarks. L.O. 24.2+
;                  $sCoverage           - [optional] a string value. Default is Null. The Document's "Coverage" Property. L.O. 24.2+
;                  $sIdentifier         - [optional] a string value. Default is Null. The Document's "Identifier" Property. L.O. 24.2+
;                  $asPublisher         - [optional] an array of strings. Default is Null. The Document's "Publisher" Property. Input must be a single dimension Array, which will overwrite any values previously set. See Remarks. L.O. 24.2+
;                  $asRelation          - [optional] an array of strings. Default is Null. The Document's "Relation" Property. Input must be a single dimension Array, which will overwrite any values previously set. See Remarks. L.O. 24.2+
;                  $sRights             - [optional] a string value. Default is Null. The Document's "Rights" Property. L.O. 24.2+
;                  $sSource             - [optional] a string value. Default is Null. The Document's "Source" Property. L.O. 24.2+
;                  $sType               - [optional] a string value. Default is Null. The Document's "Type" Property. L.O. 24.2+
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sTitle not a String.
;                  @Error 1 @Extended 3 Return 0 = $sSubject not a String.
;                  @Error 1 @Extended 4 Return 0 = $asKeywords not an Array.
;                  @Error 1 @Extended 5 Return 0 = $sComments not a String.
;                  @Error 1 @Extended 6 Return 0 = $asContributor not an Array.
;                  @Error 1 @Extended 7 Return 0 = $sCoverage not a String.
;                  @Error 1 @Extended 8 Return 0 = $sIdentifier not a String.
;                  @Error 1 @Extended 9 Return 0 = $asPublisher not an Array.
;                  @Error 1 @Extended 10 Return 0 = $asRelation not an Array.
;                  @Error 1 @Extended 11 Return 0 = $sRights not a String.
;                  @Error 1 @Extended 12 Return 0 = $sSource not a String.
;                  @Error 1 @Extended 13 Return 0 = $sType not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
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
;                  @Error 6 @Extended 1 Return 0 = Current LibreOffice version is less than 24.2.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array or a 12 Element array if current LibreOffice version 24.2 or greater. Returning array with values in order of function parameters. Any array values could be empty if no values are presently set.
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
	Local $avDescription[4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sTitle, $sSubject, $asKeywords, $sComments, $asContributor, $sCoverage, $sIdentifier, $asPublisher, $asRelation, $sRights, $sSource, $sType) Then
		If __LO_VersionCheck(24.2) Then ; These properties are only available from L.O. 24.2+.
			__LO_ArrayFill($avDescription, $oDocProp.Title(), $oDocProp.Subject(), $oDocProp.Keywords(), $oDocProp.Description(), $oDocProp.Contributor(), $oDocProp.Coverage(), _
					$oDocProp.Identifier(), $oDocProp.Publisher(), $oDocProp.Relation(), $oDocProp.Rights(), $oDocProp.Source(), $oDocProp.Type())

		Else
			__LO_ArrayFill($avDescription, $oDocProp.Title(), $oDocProp.Subject(), $oDocProp.Keywords(), $oDocProp.Description())
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sDispatch           - a string value. The Dispatch command to execute. See List of commands below.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sDispatch not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully executed dispatch command.
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
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_CursorMove
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFilePath           - a string value. Full path to save the document to, including Filename and extension. See Remarks.
;                  $bSamePath           - [optional] a boolean value. Default is False. If True, uses the path of the current document to export to. See Remarks
;                  $sFilterName         - [optional] a string value. Default is "". Filter name. If called with "" (blank string), Filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .odt extension is used instead, with the filter name of "writer8".
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
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating FilterName Property.
;                  @Error 2 @Extended 2 Return 0 = Error creating Overwrite Property.
;                  @Error 2 @Extended 3 Return 0 = Error creating Password Property.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Converting Path to/from L.O. URL
;                  @Error 3 @Extended 2 Return 0 = Document has no save path, and $bSamePath is called with True.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving FilterName.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning save path for exported document.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $atFindFormat        - [optional] an array of dll structs. Default is Null. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSrchDescriptObject not a Search Descriptor Object.
;                  @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;                  @Error 1 @Extended 5 Return 0 = $atFindFormat not an Array.
;                  @Error 1 @Extended 6 Return 0 = $atFindFormat does not contain an Object in the first Element.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Search did not return an Object, something went wrong.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Search was Successful, returning 1 dimensional array containing the objects to each match, @Extended is set to the number of matches.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $oRange              - [in/out] an object. A Range, such as a cursor with Data selected, to perform the search within.
;                  $atFindFormat        - [optional] an array of dll structs. Default is Null. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
; Return values .: Success: 1 or Array..
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;                  @Error 1 @Extended 5 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 6 Return 0 = $oRange has no data selected.
;                  @Error 1 @Extended 7 Return 0 = $atFindFormat not an Array.
;                  @Error 1 @Extended 8 Return 0 = First element in $atFindFormat not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Search did not return an Object, something went wrong.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Search was Successful, returning 1 dimensional array containing the objects for each match, @Extended is set to the number of matches.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_SearchDescriptorCreate, _LOWriter_DocFindAll, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange, _LOWriter_FindFormatModifyAlignment, _LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation, _LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak, _LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace, _LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout, _LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $atFindFormat        - [optional] an array of dll structs. Default is Null. Call with Null to skip. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
;                  $oRange              - [optional] an object. Default is Null. A Range, such as a cursor with Data selected, to perform the search within. If Null, the entire document is searched.
;                  $oLastFind           - [optional] an object. Default is Null. The last returned Object by a previous call to this function to begin the search from, if called with Null, the search begins at the start of the Document or selection, depending on if a Range is provided.
;                  $bExhaustive         - [optional] a boolean value. Default is False. If True, tests whether every result found in a document is contained in the selection or not. See remarks.
; Return values .: Success: Object or 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;                  @Error 1 @Extended 5 Return 0 = $atFindFormat not an Array.
;                  @Error 1 @Extended 6 Return 0 = First element in $atFindFormat not an Object.
;                  @Error 1 @Extended 7 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 8 Return 0 = $oRange has no data selected.
;                  @Error 1 @Extended 9 Return 0 = $oLastFind not an Object, or failed to retrieve starting position from $oRange.
;                  @Error 1 @Extended 10 Return 0 = $oLastFind incorrect Object type.
;                  @Error 1 @Extended 11 Return 0 = $bExhaustive not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Search was successful but found no matches.
;                  @Error 0 @Extended 1 Return Object = Success. Search was successful, returning the resulting Object.
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
; Name ..........: _LOWriter_DocFooterGetTextCursor
; Description ...: Create a Text cursor in a Page Style footer for text related functions.
; Syntax ........: _LOWriter_DocFooterGetTextCursor(ByRef $oPageStyle[, $bFooter = False[, $bFirstPage = False[, $bLeftPage = False[, $bRightPage = False]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $bFooter             - [optional] a boolean value. Default is False. If True, creates a text cursor for the page Footer. See Remarks.
;                  $bFirstPage          - [optional] a boolean value. Default is False. If True, creates a text cursor for the First page of the Footer. See Remarks.
;                  $bLeftPage           - [optional] a boolean value. Default is False. If True, creates a text cursor for Left pages in the Footer. See Remarks.
;                  $bRightPage          - [optional] a boolean value. Default is False. If True, creates a text cursor for right pages in the Footer. See Remarks.
; Return values .: Success: Object or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bFooter not a Boolean value.
;                  @Error 1 @Extended 3 Return 0 = $bFirstPage not a Boolean value.
;                  @Error 1 @Extended 4 Return 0 = $bLeftPage not a Boolean value.
;                  @Error 1 @Extended 5 Return 0 = $bRightPage not a Boolean value.
;                  @Error 1 @Extended 6 Return 0 = No parameters called with True.
;                  --Success--
;                  @Error 0 @Extended 0 Return Array = Success. See Remarks.
;                  @Error 0 @Extended 1 Return Object = Success. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If more than one parameter is called with True, an array is returned with the requested objects in the order that the True parameters are listed. Else the requested object is returned.
;                  If same content on left and right and first pages is active for the requested page style, you only need to use the $bFooter parameter, the others are only for when same content on first page or same content on left and right pages is deactivated.
; Related .......: _LOWriter_PageStyleGetObj, _LOWriter_PageStyleCreate, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocFooterGetTextCursor(ByRef $oPageStyle, $bFooter = False, $bFirstPage = False, $bLeftPage = False, $bRightPage = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aoReturn[1]
	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bFooter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bFirstPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bLeftPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bRightPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($bFooter = False) And ($bFirstPage = False) And ($bLeftPage = False) And ($bRightPage = False) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If $bFooter Then $aoReturn[0] = $oPageStyle.FooterText.createTextCursor()

	If $bFirstPage Then
		If IsObj($aoReturn[0]) Then ReDim $aoReturn[2]
		$aoReturn[UBound($aoReturn) - 1] = $oPageStyle.FooterTextFirst.createTextCursor()
	EndIf

	If $bLeftPage Then
		If IsObj($aoReturn[UBound($aoReturn) - 1]) Then ReDim $aoReturn[UBound($aoReturn) + 1]
		$aoReturn[UBound($aoReturn) - 1] = $oPageStyle.FooterTextLeft.createTextCursor()
	EndIf

	If $bRightPage Then
		If IsObj($aoReturn[UBound($aoReturn) - 1]) Then ReDim $aoReturn[UBound($aoReturn) + 1]
		$aoReturn[UBound($aoReturn) - 1] = $oPageStyle.FooterTextRight.createTextCursor()
	EndIf

	$vReturn = (UBound($aoReturn) = 1) ? ($aoReturn[0]) : ($aoReturn) ; If Array contains only one element, return it only outside of the array.

	Return (IsArray($vReturn)) ? (SetError($__LO_STATUS_SUCCESS, 0, $vReturn)) : (SetError($__LO_STATUS_SUCCESS, 1, $vReturn))
EndFunc   ;==>_LOWriter_DocFooterGetTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocFormSettings
; Description ...: Set or Retrieve Document Form related settings.
; Syntax ........: _LOWriter_DocFormSettings(ByRef $oDoc[, $bFormDesignMode = Null[, $bOpenInDesignMode = Null[, $bAutoControlFocus = Null[, $bUseControlWizards = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bFormDesignMode     - [optional] a boolean value. Default is Null. If True, Form design mode will be active.
;                  $bOpenInDesignMode   - [optional] a boolean value. Default is Null. If True, Form design mode will be active automatically upon opening the document.
;                  $bAutoControlFocus   - [optional] a boolean value. Default is Null. If True, the first Form control will have the focus upon opening the document.
;                  $bUseControlWizards  - [optional] a boolean value. Default is Null. If True, Control Wizards will be used.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bFormDesignMode not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bOpenInDesignMode not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bAutoControlFocus not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bUseControlWizards not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bFormDesignMode
;                  |                               2 = Error setting $bOpenInDesignMode
;                  |                               4 = Error setting $bAutoControlFocus
;                  |                               8 = Error setting $bUseControlWizards
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  In order to determine current values for $bFormDesignMode and $bUseControlWizards, a Macro is temporarily injected into the document, and subsequently deleted.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocFormSettings(ByRef $oDoc, $bFormDesignMode = Null, $bOpenInDesignMode = Null, $bAutoControlFocus = Null, $bUseControlWizards = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oStandardLibrary, $oScript
	Local $abForm[4], $abStatus[0]
	Local $aDummyArray[0]
	Local $sControlWiz = "uno:UseWizards"
	Local $sScript = "Function DocFormSettingStatus()" & @CRLF & _
			"  REM Modified from Andrew Pitonyak's Macro '5.48. Toggle design mode', found in 'Useful Macro Information', page 152, Revision: 1137." & @CRLF & _
			"  Dim oFrame            ' Current frame" & @CRLF & _
			"  Dim oDisp             ' The created dispatcher" & @CRLF & _
			"  Dim oParser           ' URL Transformer to parse the URL." & @CRLF & _
			"  Dim oStatusListener   ' The status listener that is created" & @CRLF & _
			"  Dim sListenerName     ' The type of listener that is created" & @CRLF & _
			"  Dim oUrl as New com.sun.star.util.URL" & @CRLF & _
			"  Dim oUrl2 as New com.sun.star.util.URL" & @CRLF & _
			"  Dim abStatus(2)" & @CRLF & @CRLF & _
			"  REM Parse the URL as required" & @CRLF & _
			"  oUrl.Complete = "".uno:SwitchControlDesignMode""" & @CRLF & _
			"  oUrl2.Complete = "".uno:UseWizards""" & @CRLF & _
			"  oParser = createUnoService(""com.sun.star.util.URLTransformer"")" & @CRLF & _
			"  oParser.parseStrict(oUrl)" & @CRLF & @CRLF & _
			"  oParser.parseStrict(oUrl2)" & @CRLF & @CRLF & _
			"  REM See if the current Frame supports the first UNO command" & @CRLF & _
			"  oFrame = ThisComponent.getCurrentController().getFrame()" & @CRLF & _
			"  oDisp = oFrame.queryDispatch(oUrl,"""",0)" & @CRLF & @CRLF & _
			"  REM Create the status listener" & @CRLF & _
			"  If (Not IsNull(oDisp)) Then" & @CRLF & _
			"    sListenerName = ""com.sun.star.frame.XStatusListener""" & @CRLF & _
			"    oStatusListener = CreateUnoListener(""Status_"", sListenerName)" & @CRLF & _
			"    oDisp.addStatusListener(oStatusListener, oURL)" & @CRLF & @CRLF & _
			"    abStatus(0) =  Status_Saver(Null) '" & @CRLF & _
			"    oDisp.removeStatusListener(oStatusListener, oURL)" & @CRLF & _
			"  Else" & @CRLF & _
			"    abStatus(0) = False" & @CRLF & _
			"  End If" & @CRLF & @CRLF & _
			"  REM See if the current Frame supports the second UNO command" & @CRLF & _
			"  oDisp = oFrame.queryDispatch(oUrl2,"""",0)" & @CRLF & @CRLF & _
			"  REM Create the status listener" & @CRLF & _
			"  If (Not IsNull(oDisp)) Then" & @CRLF & _
			"    sListenerName = ""com.sun.star.frame.XStatusListener""" & @CRLF & _
			"    oStatusListener = CreateUnoListener(""Status_"", sListenerName)" & @CRLF & _
			"    oDisp.addStatusListener(oStatusListener, oURL2)" & @CRLF & @CRLF & _
			"    abStatus(1) =  Status_Saver(Null) '" & @CRLF & _
			"    oDisp.removeStatusListener(oStatusListener, oURL2)" & @CRLF & _
			"  Else" & @CRLF & _
			"    abStatus(1) = False" & @CRLF & _
			"  End If" & @CRLF & _
			"  DocFormSettingStatus = abStatus" & @CRLF & @CRLF & _
			"End Function" & @CRLF & @CRLF & _
			"REM The definition of the listener requires this, but we do not use this." & @CRLF & _
			"Function Status_disposing(oEvt)" & @CRLF & _
			"End Function" & @CRLF & @CRLF & _
			"REM This is called when the status changes. In other words, when the design mode or Control Wizard is toggled and when the listener is first created." & @CRLF & _
			"Function Status_statusChanged(oEvt)" & @CRLF & _
			"  Status_Saver(oEvt.State)" & @CRLF & _
			"End Function" & @CRLF & @CRLF & _
			"Function Status_Saver(bStatus) As Boolean" & @CRLF & _
			"  Static bCurStatus As Boolean" & @CRLF & _
			"  If NOT IsNull(bStatus) Then" & @CRLF & _
			"    bCurStatus = bStatus" & @CRLF & _
			"  Else" & @CRLF & _
			"    Status_Saver = bCurStatus" & @CRLF & _
			"  End If" & @CRLF & _
			"End Function"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; Retrieving the BasicLibrary.Standard Object fails when using a newly opened document, I found a workaround by updating the following setting.
	$oDoc.BasicLibraries.VBACompatibilityMode = $oDoc.BasicLibraries.VBACompatibilityMode()

	$oStandardLibrary = $oDoc.BasicLibraries.Standard()
	If Not IsObj($oStandardLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then $oStandardLibrary.removeByName("AU3LibreOffice_UDF_Macros")

	If __LO_VarsAreNull($bFormDesignMode, $bOpenInDesignMode, $bAutoControlFocus, $bUseControlWizards) Then
		$oStandardLibrary.insertByName("AU3LibreOffice_UDF_Macros", $sScript)
		If Not $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oScript = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard.AU3LibreOffice_UDF_Macros.DocFormSettingStatus?language=Basic&location=document")
		If Not IsObj($oScript) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$abStatus = $oScript.Invoke($aDummyArray, $aDummyArray, $aDummyArray)

		$oStandardLibrary.removeByName("AU3LibreOffice_UDF_Macros")
		If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		__LO_ArrayFill($abForm, $abStatus[0], $oDoc.ApplyFormDesignMode(), $oDoc.AutomaticControlFocus(), $abStatus[1])

		Return SetError($__LO_STATUS_SUCCESS, 1, $abForm)
	EndIf

	If ($bFormDesignMode <> Null) Then
		If Not IsBool($bFormDesignMode) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.CurrentController.FormDesignMode = $bFormDesignMode

		$oStandardLibrary.insertByName("AU3LibreOffice_UDF_Macros", $sScript)
		If Not $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oScript = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard.AU3LibreOffice_UDF_Macros.DocFormSettingStatus?language=Basic&location=document")
		If Not IsObj($oScript) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$abStatus = $oScript.Invoke($aDummyArray, $aDummyArray, $aDummyArray)

		$oStandardLibrary.removeByName("AU3LibreOffice_UDF_Macros")
		If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$iError = ($abStatus[0] = $bFormDesignMode) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bOpenInDesignMode <> Null) Then
		If Not IsBool($bOpenInDesignMode) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.ApplyFormDesignMode = $bOpenInDesignMode
		$iError = ($oDoc.ApplyFormDesignMode() = $bOpenInDesignMode) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bAutoControlFocus <> Null) Then
		If Not IsBool($bAutoControlFocus) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDoc.AutomaticControlFocus = $bAutoControlFocus
		$iError = ($oDoc.AutomaticControlFocus() = $bAutoControlFocus) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bUseControlWizards <> Null) Then
		If Not IsBool($bUseControlWizards) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oStandardLibrary.insertByName("AU3LibreOffice_UDF_Macros", StringReplace($sScript, "###AUTOIT_PLACEHOLDER###", $sControlWiz))
		If Not $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oScript = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard.AU3LibreOffice_UDF_Macros.DocFormSettingStatus?language=Basic&location=document")
		If Not IsObj($oScript) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$abStatus = $oScript.Invoke($aDummyArray, $aDummyArray, $aDummyArray)

		If ($abStatus[1] <> $bUseControlWizards) Then _LOWriter_DocExecuteDispatch($oDoc, $sControlWiz) ; If the value doesn't currently match, toggle the setting.
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$abStatus = $oScript.Invoke($aDummyArray, $aDummyArray, $aDummyArray)

		$oStandardLibrary.removeByName("AU3LibreOffice_UDF_Macros")
		If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$iError = ($abStatus[1] = $bUseControlWizards) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_DocFormSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenProp
; Description ...: Set, Retrieve, or reset a Document's General Properties.
; Syntax ........: _LOWriter_DocGenProp(ByRef $oDoc[, $sNewAuthor = Null[, $iRevisions = Null[, $iEditDuration = Null[, $bApplyUserData = Null[, $bResetUserData = False]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sNewAuthor          - [optional] a string value. Default is Null. The new author of the document, can be set separately, but must be set to a string if $bResetUserData is called with True.
;                  $iRevisions          - [optional] an integer value. Default is Null. How often the document was edited and saved.
;                  $iEditDuration       - [optional] an integer value. Default is Null. The total time of editing the document (in seconds).
;                  $bApplyUserData      - [optional] a boolean value. Default is Null. If True, the user-specific settings saved within a document will be loaded with the document.
;                  $bResetUserData      - [optional] a boolean value. Default is False. If True, clears the document properties, such that it appears the document has just been created. Resets several attributes at once. See remarks.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sNewAuthor not a String and $bResetUserData called with True.
;                  @Error 1 @Extended 3 Return 0 = $sNewAuthor not a String.
;                  @Error 1 @Extended 4 Return 0 = $iRevisions not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iEditDuration not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $bApplyUserData not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error retrieving Document Settings Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sNewAuthor
;                  |                               2 = Error setting $iRevisions
;                  |                               4 = Error setting $iEditDuration
;                  |                               8 = Error setting $bApplyUserData
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 0 Return 2 = Success. Document Properties were successfully Reset.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters, except $bResetUserData, as it is not a setting.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sAuthor             - [optional] a string value. Default is Null. The initial author of the document.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sAuthor not a String.
;                  @Error 1 @Extended 3 Return 0 = $tDateStruct not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sAuthor
;                  |                               2 = Error setting $tDateStruct
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sModifiedBy         - [optional] a string value. Default is Null. The name of the last user who modified the document.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sModifiedBy not a String.
;                  @Error 1 @Extended 3 Return 0 = $tDateStruct not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sModifiedBy
;                  |                               2 = Error setting $tDateStruct
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sPrintedBy          - [optional] a string value. Default is Null. The name of the person who most recently printed the document.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sPrintedBy not a String.
;                  @Error 1 @Extended 3 Return 0 = $tDateStruct not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sPrintedBy
;                  |                               2 = Error setting $tDateStruct
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sTemplateName       - [optional] a string value. Default is Null. The name of the template from which the document was created. The value is an empty string if the document was not created from a template or if it was detached from the template
;                  $sTemplateURL        - [optional] a string value. Default is Null. The URL of the template from which the document was created. The value is an empty string if the document was not created from a template or if it was detached from the template.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sTemplateName not a String.
;                  @Error 1 @Extended 3 Return 0 = $sTemplateURL not a String.
;                  @Error 1 @Extended 4 Return 0 = $tDateStruct not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;                  @Error 3 @Extended 2 Return 0 = Error converting Computer path to Libre Office URL.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sTemplateName
;                  |                               2 = Error setting $sTemplateURL
;                  |                               4 = Error setting $tDateStruct
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1 dimension array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Document Statistics Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Array = Success. A 1 dimension, 0 based, 9 row Array of Integers, in the order described in remarks.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bReturnFull         - [optional] a boolean value. Default is False. If True, the full window title is returned, such as is used by Autoit window related functions.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bReturnFull not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the document's current Name/Title
;                  @Error 0 @Extended 1 Return String = Success. Returning the document's current Window Title, which includes the document name and usually: "-LibreOffice Writer".
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
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

	$sName = ($bReturnFull = True) ? ($oDoc.CurrentController.Frame.Title()) : ($oDoc.Title())

	Return ($bReturnFull = True) ? (SetError($__LO_STATUS_SUCCESS, 1, $sName)) : (SetError($__LO_STATUS_SUCCESS, 0, $sName))
EndFunc   ;==>_LOWriter_DocGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetPath
; Description ...: Returns a Document's current save path.
; Syntax ........: _LOWriter_DocGetPath(ByRef $oDoc[, $bReturnLibreURL = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bReturnLibreURL     - [optional] a boolean value. Default is False. If True, returns a path in Libre Office URL format, else False returns a regular Windows path.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bReturnLibreURL not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = Document has no save path.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error converting Libre URL to Computer path format.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the P.C. path to the current document's save path.
;                  @Error 0 @Extended 1 Return String = Success. Returning the Libre Office URL to the current document's save path.
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
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return ($bReturnLibreURL = True) ? (SetError($__LO_STATUS_SUCCESS, 1, $sPath)) : (SetError($__LO_STATUS_SUCCESS, 0, $sPath))
EndFunc   ;==>_LOWriter_DocGetPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetString
; Description ...: Retrieve the string of text currently selected or contained in a paragraph object.
; Syntax ........: _LOWriter_DocGetString(ByRef $oObj)
; Parameters ....: $oObj                - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions with Data selected, or a Paragraph Object returned from _LOWriter_ParObjCreateList function.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj doesn't support Paragraph Properties service.
;                  @Error 1 @Extended 3 Return 0 = $oObj is a TableCursor. Can only use View Cursor or Text Cursor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Cursor type.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. The selected text in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Libre Office documentation states that when used in Libre Basic, GetString is limited to 64kb's in size. I do not know if the same limitation applies to any outside use of GetString (such as through Autoit).
;                  If there are multiple selections, the returned value will be an empty string ("").
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGetString(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oObj.supportsService("com.sun.star.text.TextCursor") Or $oObj.supportsService("com.sun.star.text.TextViewCursor") Then
		Local $iCursorType = __LOWriter_Internal_CursorGetType($oObj)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oObj.getString())
EndFunc   ;==>_LOWriter_DocGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetViewCursor
; Description ...: Retrieve the ViewCursor Object from a Document.
; Syntax ........: _LOWriter_DocGetViewCursor(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve ViewCursor Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. The Object for the Document's View Cursor is returned for use in other Cursor related functions.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGetViewCursor(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oViewCursor

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oViewCursor = $oDoc.CurrentController.getViewCursor()

	Return (IsObj($oViewCursor)) ? (SetError($__LO_STATUS_SUCCESS, 0, $oViewCursor)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)) ; Failed to Retrieve ViewCursor
EndFunc   ;==>_LOWriter_DocGetViewCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHasPath
; Description ...: Returns whether a document has been saved to a location already or not.
; Syntax ........: _LOWriter_DocHasPath(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the document has a save location. Else False.
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

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.hasLocation())
EndFunc   ;==>_LOWriter_DocHasPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHeaderGetTextCursor
; Description ...: Create a Text cursor in a Page Style header for text related functions.
; Syntax ........: _LOWriter_DocHeaderGetTextCursor(ByRef $oPageStyle[, $bHeader = False[, $bFirstPage = False[, $bLeftPage = False[, $bRightPage = False]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or _LOWriter_PageStyleGetObj function.
;                  $bHeader             - [optional] a boolean value. Default is False. If True, creates a text cursor in the page header. See Remarks.
;                  $bFirstPage          - [optional] a boolean value. Default is False. If True, creates a text cursor in the First page of the header. See Remarks.
;                  $bLeftPage           - [optional] a boolean value. Default is False. If True, creates a text cursor in the Left pages of the header. See Remarks.
;                  $bRightPage          - [optional] a boolean value. Default is False. If True, creates a text cursor in the right pages of the header. See Remarks.
; Return values .: Success: Object or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bHeader not a Boolean value.
;                  @Error 1 @Extended 3 Return 0 = $bFirstPage not a Boolean value.
;                  @Error 1 @Extended 4 Return 0 = $bLeftPage not a Boolean value.
;                  @Error 1 @Extended 5 Return 0 = $bRightPage not a Boolean value.
;                  @Error 1 @Extended 6 Return 0 = No parameters called with True.
;                  --Success--
;                  @Error 0 @Extended 0 Return Array = Success. See Remarks.
;                  @Error 0 @Extended 1 Return Object = Success. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If more than one parameter is called with True, an array is returned with the requested objects in the order that the True parameters are listed. Else the requested object is returned.
;                  If same content on left and right and first pages is active for the requested page style, you only need to use the $bHeader parameter, the others are only for when same content on first page or same content on left and right pages is deactivated.
; Related .......: _LOWriter_PageStyleGetObj, _LOWriter_PageStyleCreate, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocHeaderGetTextCursor(ByRef $oPageStyle, $bHeader = False, $bFirstPage = False, $bLeftPage = False, $bRightPage = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aoReturn[1]
	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bHeader) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bFirstPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bLeftPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bRightPage) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($bHeader = False) And ($bFirstPage = False) And ($bLeftPage = False) And ($bRightPage = False) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If $bHeader Then $aoReturn[0] = $oPageStyle.HeaderText.createTextCursor()

	If $bFirstPage Then
		If IsObj($aoReturn[0]) Then ReDim $aoReturn[2]
		$aoReturn[UBound($aoReturn) - 1] = $oPageStyle.HeaderTextFirst.createTextCursor()
	EndIf

	If $bLeftPage Then
		If IsObj($aoReturn[UBound($aoReturn) - 1]) Then ReDim $aoReturn[UBound($aoReturn) + 1]
		$aoReturn[UBound($aoReturn) - 1] = $oPageStyle.HeaderTextLeft.createTextCursor()
	EndIf

	If $bRightPage Then
		If IsObj($aoReturn[UBound($aoReturn) - 1]) Then ReDim $aoReturn[UBound($aoReturn) + 1]
		$aoReturn[UBound($aoReturn) - 1] = $oPageStyle.HeaderTextRight.createTextCursor()
	EndIf

	$vReturn = (UBound($aoReturn) = 1) ? ($aoReturn[0]) : ($aoReturn) ; If Array contains only one element, return it only outside of the array.

	Return (IsArray($vReturn)) ? (SetError($__LO_STATUS_SUCCESS, 0, $vReturn)) : (SetError($__LO_STATUS_SUCCESS, 1, $vReturn))
EndFunc   ;==>_LOWriter_DocHeaderGetTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHyperlinkInsert
; Description ...: Insert a hyperlink into the specified document and a cursor location or other.
; Syntax ........: _LOWriter_DocHyperlinkInsert(ByRef $oDoc, ByRef $oCursor, $sLinkText, $sLinkAddress[, $bInsertAtViewCursor = False[, $bOverwrite = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions. See Remarks.
;                  $sLinkText           - a string value. Link text you want displayed (Insert the URL here too if you want the link inserted raw.)
;                  $sLinkAddress        - a string value. A URL.
;                  $bInsertAtViewCursor - [optional] a boolean value. Default is False. If True, inserts the hyperlink at the ViewCursor's position. See Remarks.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, overwrites any data selected by the $oCursor.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object, and is not called with Default keyword.
;                  @Error 1 @Extended 3 Return 0 = $sLinkText not a String.
;                  @Error 1 @Extended 4 Return 0 = $sLinkAddress not a String.
;                  @Error 1 @Extended 5 Return 0 = $bInsertAtViewCursor not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $oCursor is called with an Object, and $bInsertAtViewCursor is called with True. Change $oCursor to Default or call $bInsertAtViewCursor with False.
;                  @Error 1 @Extended 7 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $oCursor is a TableCursor, and is not supported.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Cursor Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cursor type.
;                  @Error 3 @Extended 2 Return 0 = Current ViewCursor is in unknown data type or failed detecting what data type.
;                  @Error 3 @Extended 3 Return 0 = Failed to Retrieve Text Object.
;                  --Success--
;                  @Error 0 @Extended 1 Return 1 = Success, hyperlink was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You may call this function with an already existing cursor object, which will place the Link at the cursor's current position. You can also set $oCursor to Default keyword, and set $bInsertAtViewCursor to True. This will insert the link at the current ViewCursor position. Or you can set $oCursor to Default, and leave $bInsertAtViewCursor undeclared which will insert the Link at the very end of the document.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocHyperlinkInsert(ByRef $oDoc, ByRef $oCursor, $sLinkText, $sLinkAddress, $bInsertAtViewCursor = False, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oText, $oTextCursor
	Local $iCursorType = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) And ($oCursor <> Default) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sLinkText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sLinkAddress) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bInsertAtViewCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If IsObj($oCursor) And $bInsertAtViewCursor Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If IsObj($oCursor) Or $bInsertAtViewCursor Then
		$iCursorType = (IsObj($oCursor)) ? (__LOWriter_Internal_CursorGetType($oCursor)) : (0)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		If $bInsertAtViewCursor Or ($iCursorType = $LOW_CURTYPE_VIEW_CURSOR) Then
			$oTextCursor = _LOWriter_DocCreateTextCursor($oDoc, False, True) ; create new Text cursor at ViewCursor
			If Not IsObj($oTextCursor) Or @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		EndIf

		$oTextCursor = ($iCursorType = $LOW_CURTYPE_TEXT_CURSOR) ? ($oCursor) : ($oTextCursor) ; If already was a TextCursor transfer to $oTextCursor

		$oText = __LOWriter_CursorGetText($oDoc, $oTextCursor)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		If Not IsObj($oText) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Else
		$oText = $oDoc.getText
		If Not IsObj($oText) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oTextCursor = $oText.createTextCursor()
		If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oTextCursor.gotoEnd(False)
	EndIf

	$oText.insertString($oTextCursor, $sLinkText, $bOverwrite)

	With $oTextCursor
		.goLeft(StringLen($sLinkText), False)
		.goRight(StringLen($sLinkText), True)
		.HyperLinkURL = $sLinkAddress
		.collapseToEnd()
		.goRight(1, False)
	EndWith

	Return SetError($__LO_STATUS_SUCCESS, 1, $oTextCursor)
EndFunc   ;==>_LOWriter_DocHyperlinkInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocInsertControlChar
; Description ...: Insert a control character at the cursor position.
; Syntax ........: _LOWriter_DocInsertControlChar(ByRef $oDoc, ByRef $oCursor, $iConChar[, $bOverwrite = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $iConChar            - an integer value (0-5). The control character to insert. See constants, $LOW_CON_CHAR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, and the cursor object has text selected, it is overwritten, else if False, the character is inserted to the left of the selection.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iConChar not an Integer, less than 0 or greater than 5. See Constants, $LOW_CON_CHAR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $oCursor is a TableCursor. Can only use View Cursor or Text Cursor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Cursor type.
;                  @Error 3 @Extended 2 Return 0 = Error creating Text Cursor.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Control Character was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocInsertControlChar(ByRef $oDoc, ByRef $oCursor, $iConChar, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $oTextCursor = $oCursor

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iConChar, $LOW_CON_CHAR_PAR_BREAK, $LOW_CON_CHAR_APPEND_PAR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If ($iCursorType = $LOW_CURTYPE_VIEW_CURSOR) Then $oTextCursor = _LOWriter_DocCreateTextCursor($oDoc, False, True)

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oTextCursor.Text.insertControlCharacter($oTextCursor, $iConChar, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocInsertControlChar

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocInsertString
; Description ...: Insert a string at a cursor position.
; Syntax ........: _LOWriter_DocInsertString(ByRef $oDoc, ByRef $oCursor, $sString[, $bOverwrite = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $sString             - a string value. A String to insert.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, and the cursor object has text selected, the selection is overwritten, else if False, the string is inserted to the left of the selection. If there are multiple selections, the string is inserted to the left of the last selection, and none are overwritten.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sString not a string..
;                  @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $oCursor is a TableCursor. Can only use View Cursor or Text Cursor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Cursor type.
;                  @Error 3 @Extended 2 Return 0 = Error creating Text Cursor.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. String was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To prevent accidental and unwanted newlines, @CRLF is automatically replaced with @CR to match LibreOffice's newline style.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocInsertString(ByRef $oDoc, ByRef $oCursor, $sString, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $oTextCursor = $oCursor

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If ($iCursorType = $LOW_CURTYPE_VIEW_CURSOR) Then $oTextCursor = _LOWriter_DocCreateTextCursor($oDoc, False, True)

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	; Exchange CRLF for CR to prevent errors.
	$sString = StringRegExpReplace($sString, @CRLF, @CR)

	$oTextCursor.Text.insertString($oTextCursor, $sString, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocInsertString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocIsActive
; Description ...: Tests if called document is the active document of other Libre windows.
; Syntax ........: _LOWriter_DocIsActive(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if document is the currently active Libre window. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This does NOT test if the document is the current active window in Windows, it only tests if the document is the current active document among other Libre Office documents.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocIsActive(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.CurrentController.Frame.isActive())
EndFunc   ;==>_LOWriter_DocIsActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocIsModified
; Description ...: Test whether the document has been modified since being created or since the last save.
; Syntax ........: _LOWriter_DocIsModified(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the document has been modified since last being saved.
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

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.isModified())
EndFunc   ;==>_LOWriter_DocIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocIsReadOnly
; Description ...: Tests whether a document is opened in ReadOnly mode.
; Syntax ........: _LOWriter_DocIsReadOnly(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True is document is currently Read Only, else False.
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

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.isReadOnly())
EndFunc   ;==>_LOWriter_DocIsReadOnly

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocMaximize
; Description ...: Maximize or restore a document.
; Syntax ........: _LOWriter_DocMaximize(ByRef $oDoc[, $bMaximize = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bMaximize           - [optional] a boolean value. Default is Null. If True, document window is maximized, else if False, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bMaximize not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully maximized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMaximize called with Null, returning boolean indicating if Document is currently maximized (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bMaximize is called with Null, returns a Boolean indicating if document is currently maximized (True).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocMaximize(ByRef $oDoc, $bMaximize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bMaximize) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.IsMaximized())

	If Not IsBool($bMaximize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.IsMaximized = $bMaximize

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocMaximize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocMinimize
; Description ...: Minimize or restore a document.
; Syntax ........: _LOWriter_DocMinimize(ByRef $oDoc[, $bMinimize = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bMinimize           - [optional] a boolean value. Default is Null. If True, document window is minimized, else if False, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bMinimize not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully minimized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMinimize called with Null, returning boolean indicating if Document is currently minimized (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bMinimize is called with Null, returns a Boolean indicating if document is currently minimized (True).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocMinimize(ByRef $oDoc, $bMinimize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bMinimize) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.IsMinimized())

	If Not IsBool($bMinimize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.IsMinimized = $bMinimize

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocMinimize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocOpen
; Description ...: Open an existing Writer Document, returning its object identifier.
; Syntax ........: _LOWriter_DocOpen($sFilePath[, $bConnectIfOpen = True[, $bHidden = Null[, $bReadOnly = Null[, $sPassword = Null[, $bLoadAsTemplate = Null[, $sFilterName = Null]]]]]])
; Parameters ....: $sFilePath           - a string value. Full path and filename of the file to be opened.
;                  $bConnectIfOpen      - [optional] a boolean value. Default is True(Connect). Whether to connect to the requested document if it is already open. See remarks.
;                  $bHidden             - [optional] a boolean value. Default is Null. If True, opens the document invisibly.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If True, opens the document as read-only.
;                  $sPassword           - [optional] a string value. Default is Null. The password that was used to read-protect the document, if any.
;                  $bLoadAsTemplate     - [optional] a boolean value. Default is Null. If True, opens the document as a Template, i.e. an untitled copy of the specified document is made instead of modifying the original document.
;                  $sFilterName         - [optional] a string value. Default is Null. Name of a LibreOffice filter to use to load the specified document. LibreOffice automatically selects which to use by default.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFilePath not string, or file not found.
;                  @Error 1 @Extended 2 Return 0 = Error converting filepath to URL path.
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

	If $bConnectIfOpen Then $oDoc = _LOWriter_DocConnect($sFilePath)
	If IsObj($oDoc) Then Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))

	$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOWriter_DocOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPosAndSize
; Description ...: Reposition and resize a document window.
; Syntax ........: _LOWriter_DocPosAndSize(ByRef $oDoc[, $iX = Null[, $iY = Null[, $iWidth = Null[, $iHeight = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
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

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc   ;==>_LOWriter_DocPosAndSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrint
; Description ...: Print a document using the specified settings.
; Syntax ........: _LOWriter_DocPrint(ByRef $oDoc[, $iCopies = 1[, $bCollate = True[, $vPages = "ALL"[, $bWait = True[, $iDuplexMode = $LOW_DUPLEX_OFF[, $sPrinter = ""[, $sFilePathName = ""]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iCopies             - [optional] an integer value. Default is 1. Specifies the number of copies to print.
;                  $bCollate            - [optional] a boolean value. Default is True. Advises the printer to collate the pages of the copies.
;                  $vPages              - [optional] a String or Integer value. Default is "ALL". Specifies which pages to print. See remarks.
;                  $bWait               - [optional] a boolean value. Default is True. If True, the corresponding print request will be executed synchronous. Default is to use synchronous print mode.
;                  $iDuplexMode         - [optional] an integer value (0-3). Default is $__g_iDuplexOFF. Determines the duplex mode for the print job. See Constants, $LOW_DUPLEX_* as defined in LibreOfficeWriter_Constants.au3.
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
;                  @Error 1 @Extended 7 Return 0 = $iDuplexMode not an Integer, less than 0 or greater than 3. See Constants, $LOW_DUPLEX_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $sPrinter not a String.
;                  @Error 1 @Extended 9 Return 0 = $sFilePathName not a
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "Printer Name" property.
;                  @Error 2 @Extended 2 Return 0 = Error creating "Copies" property.
;                  @Error 2 @Extended 3 Return 0 = Error creating "Collate" property.
;                  @Error 2 @Extended 4 Return 0 = Error creating "Wait" property.
;                  @Error 2 @Extended 5 Return 0 = Error creating "DuplexMode" property.
;                  @Error 2 @Extended 6 Return 0 = Error creating "Pages" property.
;                  @Error 2 @Extended 7 Return 0 = Error creating "PrintToFile" property.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error converting PrintToFile Path.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success Document was successfully printed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Based on OOoCalc UDF Print function by GMK.
;                  $vPages range can be called as entered in the user interface, as follows: "1-4,10" to print the pages 1 to 4 and 10. Default is "ALL". Must be in String format to accept more than just a single page number. e.g. This will work: "1-6,12,27" This will not 1-6,12,27. This will work: "7", This will also: 7.
;                  Setting $bWait to True is highly recommended. Otherwise following actions (as e.g. closing the Document) can fail.
; Related .......: _LO_PrintersGetNamesAlt, _LO_PrintersGetNames, _LOWriter_DocPrintSizeSettings, _LOWriter_DocPrintPageSettings, _LOWriter_DocPrintMiscSettings, _LOWriter_DocPrintIncludedSettings
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocPrint(ByRef $oDoc, $iCopies = 1, $bCollate = True, $vPages = "ALL", $bWait = True, $iDuplexMode = $LOW_DUPLEX_OFF, $sPrinter = "", $sFilePathName = "")
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
	If Not __LO_IntIsBetween($iDuplexMode, $LOW_DUPLEX_OFF, $LOW_DUPLEX_SHORT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bGraphics           - [optional] a boolean value. Default is Null. If True, the graphics contained in the document are printed.
;                  $bControls           - [optional] a boolean value. Default is Null. If True, the form control fields contained in the document are printed.
;                  $bDrawings           - [optional] a boolean value. Default is Null. If True, the drawings contained in the document are printed.
;                  $bTables             - [optional] a boolean value. Default is Null. If True, the Tables contained in the document are printed.
;                  $bHiddenText         - [optional] a boolean value. Default is Null. If True, prints text that is marked as hidden.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bGraphics not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bControls not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bDrawings not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bTables not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bHiddenText not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.DocumentSettings" Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bGraphics
;                  |                               2 = Error setting $bControls
;                  |                               4 = Error setting $bDrawings
;                  |                               8 = Error setting $bTables
;                  |                               16 = Error setting $bHiddenText
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iPaperOrient        - [optional] an integer value (0-1). Default is Null. The orientation of the paper. See Constants, $LOW_PAPER_ORIENT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPrinterName        - [optional] a string value. Default is Null. The name of the printer to send print jobs to.
;                  $iCommentsMode       - [optional] an integer value (0-3). Default is Null. If and where to print comments in the document. See Constants, $LOW_PRINT_NOTES_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bBrochure           - [optional] a boolean value. Default is Null. If True, prints the document in brochure format.
;                  $bBrochureRTL        - [optional] a boolean value. Default is Null. If True, prints the document in brochure Right to Left format.
;                  $bReversed           - [optional] a boolean value. Default is Null. If True, prints pages in reverse order.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iPaperOrient not an Integer, less than 0 or greater than 1. See Constants, $LOW_PAPER_ORIENT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $sPrinterName not a string.
;                  @Error 1 @Extended 4 Return 0 = $iCommentsMode not an Integer, less than 0 or greater than 3. See Constants, $LOW_PRINT_NOTES_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bBrochure not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bBrochureRTL not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bReversed not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.DocumentSettings" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving setting value of "CanSetPaperOrientation" from Printer.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPaperOrient
;                  |                               2 = Error setting $sPrinterName
;                  |                               4 = Error setting $iCommentsMode
;                  |                               8 = Error setting $bBrochure
;                  |                               16 = Error setting $bBrochureRTL
;                  |                               32 = Error setting $bReversed
;                  --Printer Related Errors--
;                  @Error 5 @Extended 1 Return 0 = Printer does not allow changing paper orientation.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bBlackOnly          - [optional] a boolean value. Default is Null. If True, prints all text in black only.
;                  $bLeftOnly           - [optional] a boolean value. Default is Null. If True, prints only Left(Even) pages. See remarks.
;                  $bRightOnly          - [optional] a boolean value. Default is Null. If True, prints only Right(Odd) pages. See remarks.
;                  $bBackground         - [optional] a boolean value. Default is Null. If True, prints colors and objects that are inserted to the background of the page.
;                  $bEmptyPages         - [optional] a boolean value. Default is Null. If True, automatically inserted blank pages are printed.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bBlackOnly not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bLeftOnly not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bRightOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bBackground not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bEmptyPages not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.DocumentSettings" Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bBlackOnly
;                  |                               2 = Error setting $bLeftOnly
;                  |                               4 = Error setting $bRightOnly
;                  |                               8 = Error setting $bBackground
;                  |                               16 = Error setting $bEmptyPages
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iPaperFormat        - [optional] an integer value (0-8). Default is Null. Specifies a predefined paper size or if the paper size is a user-defined size. See constants, $LOW_PAPER_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iPaperWidth         - [optional] an integer value. Default is Null. Specifies the size of the paper in Hundredths of a Millimeter (HMM). Can be a custom value or one of the constants, $LOW_PAPER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
;                  $iPaperHeight        - [optional] an integer value. Default is Null. Specifies the size of the paper in Hundredths of a Millimeter (HMM). Can be a custom value or one of the constants, $LOW_PAPER_HEIGHT_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iPaperFormat not an Integer, less than 0 or greater than 8. See constants, $LOW_PAPER_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 3 Return 0 = $iPaperWidth not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iPaperHeight not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.awt.Size" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Printer setting "CanSetPaperFormat".
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Printer setting "CanSetPaperSize".
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPaperFormat
;                  |                               2 = Error setting $iPaperWidth
;                  |                               4 = Error setting $iPaperHeight
;                  --Printer Related Errors--
;                  @Error 5 @Extended 1 Return 0 = Printer doesn't allow paper format to be set.
;                  @Error 5 @Extended 2 Return 0 = Printer doesn't allow paper size to be set.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
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
		If Not __LO_IntIsBetween($iPaperFormat, $LOW_PAPER_A3, $LOW_PAPER_USER_DEFINED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared all Redo Actions.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the document has a redo action to perform, True is returned, else False.
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

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.UndoManager.isRedoPossible())
EndFunc   ;==>_LOWriter_DocRedoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocReplaceAll
; Description ...: Replace all instances of a search.
; Syntax ........: _LOWriter_DocReplaceAll(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, $sReplaceString[, $atFindFormat = Null[, $atReplaceFormat = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a Regular Expression to Search for.
;                  $sReplaceString      - a string value. A String of text or a Regular Expression to replace any results with.
;                  $atFindFormat        - [optional] an array of dll structs. Default is Null. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
;                  $atReplaceFormat     - [optional] an array of dll structs. Default is Null. An Array of Formatting property values to replace any results with.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;                  @Error 1 @Extended 5 Return 0 = $sReplaceString not a String.
;                  @Error 1 @Extended 6 Return 0 = $atFindFormat not an Array.
;                  @Error 1 @Extended 7 Return 0 = $atReplaceFormat not an Array.
;                  @Error 1 @Extended 8 Return 0 = First Element of $atFindFormat not an Object.
;                  @Error 1 @Extended 9 Return 0 = First Element of $atReplaceFormat not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Search and Replace was successful, returning number of replacements made.
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
; Description ...: Replace all instances of a search within a selection. See Remarks.
; Syntax ........: _LOWriter_DocReplaceAllInRange(ByRef $oDoc, ByRef $oSrchDescript, ByRef $oRange, $sSearchString, $sReplaceString[, $atFindFormat = Null[, $atReplaceFormat = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $oRange              - [in/out] an object. A Range, such as a cursor with Data selected, to perform the search within.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $sReplaceString      - a string value. A String of text or a regular expression to replace any results with.
;                  $atFindFormat        - [optional] an array of dll structs. Default is Null. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
;                  $atReplaceFormat     - [optional] an array of dll structs. Default is Null. An Array of Formatting property values to replace any results with.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;                  @Error 1 @Extended 4 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 5 Return 0 = $oRange contains no selected Data.
;                  @Error 1 @Extended 6 Return 0 = $sSearchString not a String.
;                  @Error 1 @Extended 7 Return 0 = $sReplaceString not a String.
;                  @Error 1 @Extended 8 Return 0 = $atFindFormat not an Array.
;                  @Error 1 @Extended 9 Return 0 = $atReplaceFormat not an Array.
;                  @Error 1 @Extended 10 Return 0 = First Element in $atFindFormat not a Property Object.
;                  @Error 1 @Extended 11 Return 0 = First Element in $atReplaceFormat not a Property Object.
;                  @Error 1 @Extended 12 Return 0 = Paragraph Style Name called in $sReplaceString does not exist.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating backup of ViewCursor location and selection.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 3 Return 0 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Finding all results in Range.
;                  @Error 3 @Extended 2 Return 0 = Error searching for property values.
;                  @Error 3 @Extended 3 Return 0 = Error finding temporary property to use.
;                  @Error 3 @Extended 4 Return 0 = Error retrieving current selection and ViewCursor position.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Search and Replace was successful, returning number of replacements.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Libre Office does not offer a method to replace only results within a selection, consequently I have had to create my own. This function sometimes uses the "FindAllInRange" function, so any errors with Find/Replace formatting causing deletions will cause problems here. As best as I can tell all options for find and replace should be available, Formatting, Paragraph styles etc.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Document is ReadOnly or Document has no save location, try SaveAs.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Document Successfully saved.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFilePath           - a string value. Full path to save the document to, including Filename and extension.
;                  $sFilterName         - [optional] a string value. Default is "". The filter name. Calling "" (blank string), means the filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .odt extension is used instead, with the filter name of "writer8".
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
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating FilterName Property
;                  @Error 2 @Extended 2 Return 0 = Error creating Overwrite Property
;                  @Error 2 @Extended 3 Return 0 = Error creating Password Property
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Converting Path to/from L.O. URL
;                  @Error 3 @Extended 2 Return 0 = Error retrieving FilterName.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Successfully Saved the document. Returning document save path.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObj                - [in/out] an object. Default is Null. A selectable object. A Text Cursor with text selected, A ViewCursor with Text Selected, a Table Cursor with cells selected, a Shape or Frame Object, etc.
;                  $bReturnMultiAsObj   - [optional] a boolean value. Default is False. If True, when Multiple selections are present, they will be returned as a single Object. See Remarks.
; Return values .: Success: 1, Object or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bReturnMultiAsObj not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a TextCursor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current selection.
;                  @Error 3 @Extended 2 Return 0 =There is no text selected.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve count of multiple selections.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify current selection.
;                  @Error 3 @Extended 5 Return 0 = Failed to select object.
;                  --Success--
;                  @Error 0 @Extended -6 Return Object = Success. The current selection is within a Table, returning a Table Cursor Object.
;                  @Error 0 @Extended -5 Return Object = Success. The current selection was a Frame, returning a Frame Object.
;                  @Error 0 @Extended -4 Return Object = Success. The current selection was a Shape, returning a Shape Object.
;                  @Error 0 @Extended -3 Return Object = Success. The current selection was a Chart or other OLE object, returning the Object.
;                  @Error 0 @Extended -2 Return Object = Success. The current selection was an Image, returning a Image Object.
;                  @Error 0 @Extended -1 Return Object = Success. The current selection is multiple disconnected selections and $bReturnMultiAsObj was True, Returning a single Object.
;                  @Error 0 @Extended 0 Return 1 = Success. Object called in $oObj successfully selected.
;                  @Error 0 @Extended 1 Return Object = Success. The current selection is a single span of text. Returning a Text Cursor.
;                  @Error 0 @Extended ? Return Array = Success. The current selection is multiple disconnected selections. Returning an Array of Text Cursors. @Extended is set to number of results.
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

	Return ($bSelect) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0))
EndFunc   ;==>_LOWriter_DocSelection

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocToFront
; Description ...: Bring the called document to the front of the other windows.
; Syntax ........: _LOWriter_DocToFront(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully ended the last Undo Action group recording.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared all Undo and Redo Actions.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the document has an undo action to perform, True is returned, else False.
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

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.UndoManager.isUndoPossible())
EndFunc   ;==>_LOWriter_DocUndoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoReset
; Description ...: Reset the UndoManager.
; Syntax ........: _LOWriter_DocUndoReset(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully reset the undo manager.
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
; Name ..........: _LOWriter_DocViewCursorGetPosition
; Description ...: Retrieve View Cursor position in Hundredths of a Millimeter (HMM).
; Syntax ........: _LOWriter_DocViewCursorGetPosition(ByRef $oCursor)
; Parameters ....: $oCursor             - [in/out] an object. A View Cursor Object returned by _LOWriter_DocGetViewCursor function.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not a View Cursor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error determining Cursor type.
;                  --Success--
;                  @Error 0 @Extended ? Return Integer = Success. Current Cursor Coordinate position relative to the top-left of the first page of the document is returned. @Extended is the "X" coordinate, and Return value is the "Y" Coordinate. In Hundredths of a Millimeter (HMM).
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_CursorMove, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocViewCursorGetPosition(ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType <> $LOW_CURTYPE_VIEW_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, $oCursor.getPosition().X(), $oCursor.getPosition().Y())
EndFunc   ;==>_LOWriter_DocViewCursorGetPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocVisible
; Description ...: Set or retrieve the current visibility of a document.
; Syntax ........: _LOWriter_DocVisible(ByRef $oDoc[, $bVisible = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the document is visible.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
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
Func _LOWriter_DocVisible(ByRef $oDoc, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bVisible) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.isVisible())

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.Visible = $bVisible
	$iError = ($oDoc.CurrentController.Frame.ContainerWindow.isVisible() = $bVisible) ? (0) : (1)

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOWriter_DocVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocZoom
; Description ...: Modify the zoom value for a document.
; Syntax ........: _LOWriter_DocZoom(ByRef $oDoc[, $iZoom = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iZoom               - [optional] an integer value (20-600). Default is Null. The zoom percentage.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iZoom not an Integer, less than 20 or greater than 600.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iZoom
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = $iZoom set successfully.
;                  @Error 0 @Extended 1 Return Integer =  All optional parameters were called with Null, returning current zoom value.
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

	Local $iError = 0
	Local $oServiceManager, $oDispatcher
	Local $aArgs[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iZoom) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.CurrentController.ViewSettings.ZoomValue())

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
