#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

#include "LibreOfficeWriter_Par.au3"

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
; _LOWriter_DocBookmarkDelete
; _LOWriter_DocBookmarkGetAnchor
; _LOWriter_DocBookmarkGetObj
; _LOWriter_DocBookmarkInsert
; _LOWriter_DocBookmarkModify
; _LOWriter_DocBookmarksHasName
; _LOWriter_DocBookmarksList
; _LOWriter_DocClose
; _LOWriter_DocConnect
; _LOWriter_DocConvertTableToText
; _LOWriter_DocConvertTextToTable
; _LOWriter_DocCreate
; _LOWriter_DocCreateTextCursor
; _LOWriter_DocDescription
; _LOWriter_DocEnumPrinters
; _LOWriter_DocEnumPrintersAlt
; _LOWriter_DocExecuteDispatch
; _LOWriter_DocExport
; _LOWriter_DocFindAll
; _LOWriter_DocFindAllInRange
; _LOWriter_DocFindNext
; _LOWriter_DocFooterGetTextCursor
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
; _LOWriter_DocHasFrameName
; _LOWriter_DocHasImageName
; _LOWriter_DocHasPath
; _LOWriter_DocHasTableName
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
; _LOWriter_DocRedoCurActionTitle
; _LOWriter_DocRedoGetAllActionTitles
; _LOWriter_DocRedoIsPossible
; _LOWriter_DocReplaceAll
; _LOWriter_DocReplaceAllInRange
; _LOWriter_DocSave
; _LOWriter_DocSaveAs
; _LOWriter_DocToFront
; _LOWriter_DocUndo
; _LOWriter_DocUndoCurActionTitle
; _LOWriter_DocUndoGetAllActionTitles
; _LOWriter_DocUndoIsPossible
; _LOWriter_DocViewCursorGetPosition
; _LOWriter_DocVisible
; _LOWriter_DocZoom
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkDelete
; Description ...: Selete a Bookmark.
; Syntax ........: _LOWriter_DocBookmarkDelete(Byref $oDoc, Byref $oBookmark)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oBookmark           - [in/out] an object. A Bookmark Object from a previous _LOWriter_DocBookmarkInsert, or _LOWriter_DocBookmarkGetObj function to delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oBookmark not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to delete Bookmark, but document still contains a Bookmark by that name.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested Bookmark.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocBookmarkInsert, _LOWriter_DocBookmarkGetObj, _LOWriter_DocBookmarksList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkDelete(ByRef $oDoc, ByRef $oBookmark)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sBookmarkName

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oBookmark) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$sBookmarkName = $oBookmark.Name()

	$oBookmark.dispose()

	Return (_LOWriter_DocBookmarksHasName($oDoc, $sBookmarkName)) ? SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocBookmarkDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkGetAnchor
; Description ...: Retrieve a Bookmark's Anchor cursor Object.
; Syntax ........: _LOWriter_DocBookmarkGetAnchor(Byref $oBookmark)
; Parameters ....: $oBookmark           - [in/out] an object. A Bookmark Object from a previous _LOWriter_DocBookmarkInsert, or _LOWriter_DocBookmarkGetObj function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oBookmark not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Bookmark anchor Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning requested Bookmark Anchor Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Anchor cursor returned is just a Text Cursor placed at the anchor's position.
; Related .......: _LOWriter_DocBookmarkGetObj, _LOWriter_DocBookmarkInsert, _LOWriter_CursorMove, _LOWriter_DocGetString,
;					_LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkGetAnchor(ByRef $oBookmark)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBookAnchor

	If Not IsObj($oBookmark) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oBookAnchor = $oBookmark.Anchor.Text.createTextCursorByRange($oBookmark.Anchor())
	If Not IsObj($oBookAnchor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oBookAnchor)
EndFunc   ;==>_LOWriter_DocBookmarkGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkGetObj
; Description ...: Retrieve a Bookmark Object by name.
; Syntax ........: _LOWriter_DocBookmarkGetObj(Byref $oDoc, $sBookmarkName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sBookmarkName       - a string value. The Bookmark name to retrieve the Object for.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sBookmarkName not an Object.
;				   @Error 1 @Extended 3 Return 0 = Document does not contain a Bookmark named as called in $sBookmarkName.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve requested Bookmark Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully retrieved requested Bookmark Object. Returning requested Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocBookmarksList, _LOWriter_DocBookmarkModify, _LOWriter_DocBookmarkDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkGetObj(ByRef $oDoc, $sBookmarkName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBookmark

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sBookmarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If Not _LOWriter_DocBookmarksHasName($oDoc, $sBookmarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$oBookmark = $oDoc.Bookmarks.getByName($sBookmarkName)
	If Not IsObj($oBookmark) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oBookmark)
EndFunc   ;==>_LOWriter_DocBookmarkGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkInsert
; Description ...: Insert a Bookmark into a document.
; Syntax ........: _LOWriter_DocBookmarkInsert(Byref $oDoc, Byref $oCursor[, $bOverwrite = False[, $sBookmarkName = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten.
;				   +						If False, content will be inserted to the left of any selection.
;                  $sBookmarkName       - [optional] a string value. Default is Null. The Name of the Bookmark to create. See Remarks.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sBookmarkName not a String.
;				   @Error 1 @Extended 6 Return 0 = $sBookmarkName contains illegal characters, /\@:*?";,.# .
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.Bookmark" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Bookmark into the document. Returning the Bookmark Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the cursor used to insert a Bookmark has text selected, the Bookmark will envelope the text, else the Bookmark will be inserted at a single point.
;					A Bookmark name cannot contain the following characters: / \ @ : * ? " ; , . #
;					If the document already contains a Bookmark by the same name, Libre Office adds a digit after the name, such as Bookmark 1, Bookmark 2 etc.
; Related .......: _LOWriter_DocBookmarkModify, _LOWriter_DocBookmarkDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sBookmarkName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBookmark

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$oBookmark = $oDoc.createInstance("com.sun.star.text.Bookmark")
	If Not IsObj($oBookmark) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($sBookmarkName <> Null) Then
		If Not IsString($sBookmarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		If StringRegExp($sBookmarkName, '[/\@:*?";,.#]') Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0) ; Invalid Characters in Name.
		$oBookmark.Name = $sBookmarkName
	Else
		$oBookmark.Name = "Bookmark "
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oBookmark, $bOverwrite)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oBookmark)
EndFunc   ;==>_LOWriter_DocBookmarkInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarkModify
; Description ...: Set or Retrieve a Bookmark's settings.
; Syntax ........: _LOWriter_DocBookmarkModify(Byref $oDoc, Byref $oBookmark[, $sBookmarkName = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oBookmark           - [in/out] an object. A Bookmark Object from a previous _LOWriter_DocBookmarkInsert, or _LOWriter_DocBookmarkGetObj function.
;                  $sBookmarkName       - [optional] a string value. Default is Null. The new name to name the bookmark.
; Return values .: Success: 1 or String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oBookmark not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sBookmarkName not a String.
;				   @Error 1 @Extended 4 Return 0 = $sBookmarkName contains illegal characters, /\@:*?";,.# .
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 	0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |							1 = Error setting $sBookmarkName
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Bookmark name successfully modified.
;				   @Error 0 @Extended 0 Return String = Success. $sBookmarkName set to Null, returning current Bookmark name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					A Bookmark name cannot contain the following characters: / \ @ : * ? " ; , . #
;					If the document already contains a Bookmark by the same name, Libre Office adds a digit after the name, such as Bookmark 1, Bookmark 2 etc.
; Related .......: _LOWriter_DocBookmarkGetObj, _LOWriter_DocBookmarkInsert, _LOWriter_DocBookmarkDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarkModify(ByRef $oDoc, ByRef $oBookmark, $sBookmarkName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oBookmark) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sBookmarkName) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oBookmark.Name())

	If Not IsString($sBookmarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If StringRegExp($sBookmarkName, '[/\@:*?";,.#]') Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0) ; Invalid Characters in Name.
	$oBookmark.Name = $sBookmarkName
	$iError = ($oBookmark.Name() = $sBookmarkName) ? $iError : BitOR($iError, 1)

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocBookmarkModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarksHasName
; Description ...: Check if a document contains a Bookmark by name.
; Syntax ........: _LOWriter_DocBookmarksHasName(Byref $oDoc, $sBookmarkName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sBookmarkName       - a string value. The Bookmark name to search for.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sBookmarkName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Bookmarks Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. If the document contains a Bookmark by the called name, then True is returned, Else false.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarksHasName(ByRef $oDoc, $sBookmarkName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBookmarks

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sBookmarkName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oBookmarks = $oDoc.getBookmarks()
	If Not IsObj($oBookmarks) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oBookmarks.hasByName($sBookmarkName))
EndFunc   ;==>_LOWriter_DocBookmarksHasName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocBookmarksList
; Description ...: Retrieve an Array of Bookmark names.
; Syntax ........: _LOWriter_DocBookmarksList(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Array of Bookmark Names.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully searched for Bookmarks, but document does not contain any.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for Bookmarks, returning Array of Bookmark Names, with @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocBookmarkGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocBookmarksList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asBookmarkNames[0]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$asBookmarkNames = $oDoc.Bookmarks.getElementNames()
	If Not IsArray($asBookmarkNames) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	Return (UBound($asBookmarkNames) > 0) ? SetError($__LOW_STATUS_SUCCESS, UBound($asBookmarkNames), $asBookmarkNames) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocBookmarksList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocClose
; Description ...: Close an existing Writer Document, returning its save path if applicable.
; Syntax ........: _LOWriter_DocClose(Byref $oDoc[, $bSaveChanges = True[, $sSaveName = ""[, $bDeliverOwnership = True]]])
; Parameters ....: $oDoc           		- [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bSaveChanges        - [optional] a boolean value. Default is True. If true, saves changes if any were made before closing. See remarks.
;                  $sSaveName           - [optional] a string value. Default is "". The file name to save the file as if the file hasn't been saved before. See Remarks.
;                  $bDeliverOwnership   - [optional] a boolean value. Default is True. If True, deliver ownership of the document Object from the script to LibreOffice, recommended is True.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bSaveChanges not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sSaveName not a String.
;				   @Error 1 @Extended 4 Return 0 = $bDeliverOwnership not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Path Conversion to L.O. URL Failed.
;				   @Error 3 @Extended 2 Return 0 = Error while retrieving FilterName.
;				   @Error 3 @Extended 3 Return 0 = Error While setting Filter Name properties.
;				   --Success--
;				   @Error 0 @Extended 1 Return String = Success, document was successfully closed, and was saved to the returned file Path.
;				   @Error 0 @Extended 2 Return String = Success, Document was successfully closed, document's changes were saved to its existing location.
;				   @Error 0 @Extended 3 Return String = Success, Document was successfully closed, document either had no changes to save, or $bSaveChanges was set to False.
;				   +			If document had a save location, or if document was saved to a location, it is returned, else an empty string is returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bSaveChanges is true and the document hasn't been saved yet, the document is saved to the desktop,
;					If $sSaveName is undefined, it is saved as an .odt document to the desktop, named
;					Year-Month-Day_Hour-Minute-Second.odt. $sSaveName may be just a name, in which case the file will be saved
;					in .odt format. Or you may define your own format by including an extension, such as "Test.docx"
; Related .......: _LOWriter_DocOpen, _LOWriter_DocConnect, _LOWriter_DocCreate, _LOWriter_DocSaveAs, _LOWriter_DocSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocClose(ByRef $oDoc, $bSaveChanges = True, $sSaveName = "", $bDeliverOwnership = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sDocPath = "", $sSavePath, $sFilterName
	Local $aArgs[1]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bSaveChanges) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSaveName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDeliverOwnership) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	If Not $oDoc.hasLocation() And ($bSaveChanges = True) Then
		$sSavePath = @DesktopDir & "\"
		If ($sSaveName = "") Or ($sSaveName = " ") Then
			$sSaveName = @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "-" & @MIN & "-" & @SEC & ".odt"
			$sFilterName = "writer8"
		EndIf

		$sSavePath = _LOWriter_PathConvert($sSavePath & $sSaveName, 1)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		If $sFilterName = "" Then $sFilterName = __LOWriter_FilterNameGet($sSavePath)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
		$aArgs[0] = __LOWriter_SetPropertyValue("FilterName", $sFilterName)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 3, 0)

	EndIf

	If ($bSaveChanges = True) Then

		If $oDoc.hasLocation() Then
			$oDoc.store()
			$sDocPath = _LOWriter_PathConvert($oDoc.getURL(), $LOW_PATHCONV_PCPATH_RETURN)
			$oDoc.Close($bDeliverOwnership)
			Return SetError($__LOW_STATUS_SUCCESS, 2, $sDocPath)
		Else
			$oDoc.storeAsURL($sSavePath, $aArgs)
			$oDoc.Close($bDeliverOwnership)
			Return SetError($__LOW_STATUS_SUCCESS, 1, _LOWriter_PathConvert($sSavePath, $LOW_PATHCONV_PCPATH_RETURN))
		EndIf

	EndIf

	If $oDoc.hasLocation() Then $sDocPath = _LOWriter_PathConvert($oDoc.getURL(), $LOW_PATHCONV_PCPATH_RETURN)
	$oDoc.Close($bDeliverOwnership)
	Return SetError($__LOW_STATUS_SUCCESS, 3, $sDocPath)

EndFunc   ;==>_LOWriter_DocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocConnect
; Description ...: Connect to an already opened instance of a specified LibreOffice document.
; Syntax ........: _LOWriter_DocConnect($sFile[, $bConnectCurrent = False[, $bConnectAll = False]])
; Parameters ....: $sFile               - a string value. A Full or partial file path, or a full or partial file name. See remarks. Can be an empty string is $bConnectAll or $bConnectCurrent is True.
;                  $bConnectCurrent     - [optional] a boolean value. Default is False. If True, returns the currently active, or last active Document, unless it is not a Text Document.
;                  $bConnectAll         - [optional] a boolean value. Default is False. If True, returns an array containing all open Libre Text Documents. See remarks.
; -Return values .: Success: Object or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sFile not a string.
;				   @Error 1 @Extended 2 Return 0 = $bConnectCurrent Not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bConnectAll Not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating ServiceManager object.
;				   @Error 2 @Extended 2 Return 0 = Error creating Desktop object.
;				   @Error 2 @Extended 3 Return 0 = Error creating enumeration of open documents.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error converting path to Libre Office URL.
;				   --Document Errors--
;				   @Error 5 @Extended 1 Return 0 = No matches found.
;				   @Error 5 @Extended 2 Return 0 = Current Component not a Text Document.
;				   @Error 5 @Extended 3 Return 0 = No open Libre Office documents found.
;				   --Success--
;				   @Error 0 @Extended 1 Return Object = The Object for the current, or last active document is returned.
;				   @Error 0 @Extended 2 Returns Array = An Array of all open Libre Text documents is returned. See remarks.
;				   @Error 0 @Extended 3 Return Object = The Object for the document with matching URL is returned.
;				   @Error 0 @Extended 4 Return Object = The Object for the document with matching Title is returned.
;				   @Error 0 @Extended 5 Return Object = A partial Title or Path search found only one match, returning the Object for the found document.
;				   @Error 0 @Extended 6 Return Array = An Array of all matching Libre Text documents from a partial Title or Path search. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sFile can be either the full Path (Name and extension included; i.e:
;					C:\file\Test.odt Or file:///C:/file/Test.odt) of the document,
;					or the full Title with extension, (i.e: Test.odt),
;					or a partial file path (i.e: file1\file2\Test Or file1\file2 Or file1/file2/ etc.),
;					or a partial name (i.e: test, would match Test1.odt, Test2.docx etc.).
;					Partial file path searches and file name searches, as well as the connect all option, return arrays with three columns per result.
;					($aArray[0][3]. each result is stored in a separate row;
;				    Row 1, Column 0 contain the Object variable for that document.
;						$aArray[0][0] = $oDoc
;					Row 1, Column 1 contains the Document's full title and extension.
;				    	$aArray[0][1] = This Test File.Docx
;					Row 1, Column 2 contains the document's full file path.
;				    	$aArray[0][2] = C:\Folder1\Folder2\This Test File.Docx
;					Row 2, Column 0 contain the Object variable for the next document. And so on.
;					    $aArray[1][0] = $oDoc2
; Related .......: _LOWriter_DocOpen, _LOWriter_DocClose, _LOWriter_DocCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocConnect($sFile, $bConnectCurrent = False, $bConnectAll = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local Const $STR_STRIPLEADING = 1
	Local $aoConnectAll[1], $aoPartNameSearch[1]
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop

	If Not IsString($sFile) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bConnectCurrent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bConnectAll) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	If Not $oDesktop.getComponents.hasElements() Then Return SetError($__LOW_STATUS_DOC_ERROR, 3, 0) ; no L.O open
	$oEnumDoc = $oDesktop.getComponents.createEnumeration()
	If Not IsObj($oEnumDoc) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	If $bConnectCurrent Then
		$oDoc = $oDesktop.currentComponent()
		Return ($oDoc.supportsService("com.sun.star.text.TextDocument")) ? SetError($__LOW_STATUS_SUCCESS, 1, $oDoc) : SetError($__LOW_STATUS_DOC_ERROR, 2, 0)
	EndIf

	If $bConnectAll Then

		ReDim $aoConnectAll[1][3]
		$iCount = 0
		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService("com.sun.star.text.TextDocument") Then

				ReDim $aoConnectAll[$iCount + 1][3]
				$aoConnectAll[$iCount][0] = $oDoc
				$aoConnectAll[$iCount][1] = $oDoc.Title()
				$aoConnectAll[$iCount][2] = _LOWriter_PathConvert($oDoc.getURL(), $LOW_PATHCONV_PCPATH_RETURN)
				$iCount += 1
			EndIf
			Sleep(10)
		WEnd
		Return SetError($__LOW_STATUS_SUCCESS, 2, $aoConnectAll)

	EndIf

	$sFile = StringStripWS($sFile, $STR_STRIPLEADING)
	If StringInStr($sFile, "\") Then $sFile = _LOWriter_PathConvert($sFile, $LOW_PATHCONV_OFFICE_RETURN) ; Convert to L.O File path.
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	If StringInStr($sFile, "file:///") Then ; URL/Path and Name search

		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()

			If ($oDoc.getURL() == $sFile) Then Return SetError($__LOW_STATUS_SUCCESS, 3, $oDoc) ; Match
		WEnd
		Return SetError($__LOW_STATUS_DOC_ERROR, 1, 0) ; no match

	Else
		If Not StringInStr($sFile, "/") And StringInStr($sFile, ".") Then ; Name with extension only search
			While $oEnumDoc.hasMoreElements()
				$oDoc = $oEnumDoc.nextElement()
				If StringInStr($oDoc.Title, $sFile) Then Return SetError($__LOW_STATUS_SUCCESS, 4, $oDoc) ; Match
			WEnd
			Return SetError($__LOW_STATUS_DOC_ERROR, 1, 0) ; no match
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
					$aoPartNameSearch[$iCount][2] = _LOWriter_PathConvert($oDoc.getURL, $LOW_PATHCONV_PCPATH_RETURN)
					$iCount += 1
				EndIf
			Else
				If StringInStr($oDoc.Title, $sFile) Then
					ReDim $aoPartNameSearch[$iCount + 1][3]
					$aoPartNameSearch[$iCount][0] = $oDoc
					$aoPartNameSearch[$iCount][1] = $oDoc.Title
					$aoPartNameSearch[$iCount][2] = _LOWriter_PathConvert($oDoc.getURL, $LOW_PATHCONV_PCPATH_RETURN)
					$iCount += 1
				EndIf
			EndIf

		WEnd
		If IsString($aoPartNameSearch[0][1]) Then
			If (UBound($aoPartNameSearch) = 1) Then
				Return SetError($__LOW_STATUS_SUCCESS, 5, $aoPartNameSearch[0][0]) ; matches
			Else
				Return SetError($__LOW_STATUS_SUCCESS, 6, $aoPartNameSearch) ; matches
			EndIf

		Else
			Return SetError($__LOW_STATUS_DOC_ERROR, 1, 0) ; no match
		EndIf

	EndIf

EndFunc   ;==>_LOWriter_DocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocConvertTableToText
; Description ...: Convert a Table to Text, separated by a delimiter.
; Syntax ........: _LOWriter_DocConvertTableToText(Byref $oDoc, Byref $oTable, $sDelimiter)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oTable              - [in/out] an object. A Table Object returned from _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName functions.
;                  $sDelimiter          - [optional] a string value. Default is @TAB. A character to separate each column by, such as a Tab etc.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oTable not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sDelimiter not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve ViewCursor object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create a backup of the ViewCursor's current location.
;				   @Error 2 @Extended 3 Return 0 = Failed to create a Text Cursor in the first cell.
;				   @Error 2 @Extended 4 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;				   @Error 2 @Extended 5 Return 0 = Failed to create "com.sun.star.frame.DispatchHelper" Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve array of CellNames.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Table was successfully converted to text.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function temporarily moves the Viewcursor to the Table indicated, and then attempts to restore the ViewCursor to its former position.
;					This could cause a COM error if the Cursor was presently in the Table.
; Related .......: _LOWriter_DocConvertTextToTable, _LOWriter_TableGetObjByName, _LOWriter_TableGetObjByCursor,
;					_LOWriter_TableInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocConvertTableToText(ByRef $oDoc, ByRef $oTable, $sDelimiter = @TAB)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aArgs[1]
	Local $asCellNames
	Local $oServiceManager, $oDispatcher, $oCellTextCursor, $oViewCursor, $oViewCursorBackup

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTable) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sDelimiter) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$aArgs[0] = __LOWriter_SetPropertyValue("Delimiter", $sDelimiter)

	$asCellNames = $oTable.getCellNames()
	If Not IsArray($asCellNames) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	; Retrieve the ViewCursor.
	$oViewCursor = $oDoc.CurrentController.getViewCursor()
	If Not IsObj($oViewCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	; Create a Text cursor at the current viewCursor position to move the Viewcursor back to.
	$oViewCursorBackup = _LOWriter_DocCreateTextCursor($oDoc, False, True)
	If Not IsObj($oViewCursorBackup) Then
		$oViewCursorBackup = _LOWriter_DocCreateTextCursor($oDoc, False) ; If That Failed, create a Backup Cursor at the beginning of the document.
		If Not IsObj($oViewCursorBackup) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	EndIf

	; Retrieve the first cell  in the table and create a text cursor in it to move the ViewCursor to.
	$oCellTextCursor = $oTable.getCellByName($asCellNames[0]).Text.createTextCursor()
	If Not IsObj($oCellTextCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	$oViewCursor.gotoRange($oCellTextCursor, False)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 4, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LOW_STATUS_INIT_ERROR, 5, 0)

	$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:ConvertTableToText", "", 0, $aArgs)

	; Restore the ViewCursor to its previous location.
	$oViewCursor.gotoRange($oViewCursorBackup, False)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)

EndFunc   ;==>_LOWriter_DocConvertTableToText

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocConvertTextToTable
; Description ...: Convert some selected text into a Table.
; Syntax ........: _LOWriter_DocConvertTextToTable(Byref $oDoc, Byref $oCursor[, $sDelimiter = @TAB[, $bHeader = False[, $iRepeatHeaderLines = 0[, $bBorder = False[, $bDontSplitTable = False]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions. Default is Null. See Remarks.
;                  $sDelimiter          - [optional] a string value. Default is @TAB. A character to the text into each column by, such as a Tab etc.
;                  $bHeader             - [optional] a boolean value. Default is False. If True, Formats the first row of the new table as a heading.
;                  $iRepeatHeaderLines  - [optional] an integer value. Default is 0. If greater than 0, then Repeats the first n rows as a header.
;                  $bBorder             - [optional] a boolean value. Default is False. If True, Adds a border to the table and the table cells.
;                  $bDontSplitTable     - [optional] a boolean value. Default is False. If True, Does not divide the table across pages.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sDelimiter not a String.
;				   @Error 1 @Extended 4 Return 0 = $bHeader not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iRepeatHeaderLines not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $bBorder not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $bDontSplitTable not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $oCursor is a Table Cursor and cannot be used.
;				   @Error 1 @Extended 9 Return 0 = $oCursor has no data selected.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve TextTables Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to create "com.sun.star.frame.DispatchHelper" Object.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve ViewCursor object.
;				   @Error 2 @Extended 5 Return 0 = Failed to create a backup of the ViewCursor's current location.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to $oCursor's cursor type.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve new Table's Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Text was successfully converted to a Table, returning the new Table's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function temporarily moves the Viewcursor to and selectes the Text, and then attempts to restore the ViewCursor to its former position.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					 _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					 _LOWriter_DocFooterGetTextCursor, _LOWriter_ParObjCreateList, _LOWriter_DocConvertTableToText
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocConvertTextToTable(ByRef $oDoc, ByRef $oCursor, $sDelimiter = @TAB, $bHeader = False, $iRepeatHeaderLines = 0, $bBorder = False, $bDontSplitTable = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asTables[0]
	Local $atArgs[5]
	Local $oServiceManager, $oDispatcher, $oViewCursor, $oViewCursorBackup, $oTables, $oTable
	Local $iCursorType

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sDelimiter) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bHeader) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iRepeatHeaderLines) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bBorder) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bDontSplitTable) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)

	$oTables = $oDoc.TextTables()
	If Not IsObj($oTables) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	ReDim $asTables[$oTables.getCount()]
	; Store all current Table Names.
	For $i = 0 To $oTables.getCount() - 1
		$asTables[$i] = $oTables.getByIndex($i).Name()
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))     ;Sleep every x cycles.
	Next

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)

	; If Cursor has no data selected, return error.
	If $oCursor.isCollapsed() Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	$atArgs[0] = __LOWriter_SetPropertyValue("Delimiter", $sDelimiter)
	$atArgs[1] = __LOWriter_SetPropertyValue("WithHeader", $bHeader)
	$atArgs[2] = __LOWriter_SetPropertyValue("RepeatHeaderLines", $iRepeatHeaderLines)
	$atArgs[3] = __LOWriter_SetPropertyValue("WithBorder", $bBorder)
	$atArgs[4] = __LOWriter_SetPropertyValue("DontSplitTable", $bDontSplitTable)

	If ($iCursorType = $LOW_CURTYPE_TEXT_CURSOR) Then

		; Retrieve the ViewCursor.
		$oViewCursor = $oDoc.CurrentController.getViewCursor()
		If Not IsObj($oViewCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 4, 0)

		; Create a Text cursor at the current viewCursor position to move the Viewcursor back to.
		$oViewCursorBackup = _LOWriter_DocCreateTextCursor($oDoc, False, True)
		If Not IsObj($oViewCursorBackup) Then Return SetError($__LOW_STATUS_INIT_ERROR, 5, 0)

		$oViewCursor.gotoRange($oCursor, False)

		$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:ConvertTextToTable", "", 0, $atArgs)

		; Restore the ViewCursor to its previous location.
		$oViewCursor.gotoRange($oViewCursorBackup, False)
	Else

		$oDispatcher.executeDispatch($oDoc.CurrentController(), ".uno:ConvertTextToTable", "", 0, $atArgs)
	EndIf

	; Obtain the newly created table object by comparing the original table names to the new list of tables.
	; If none match, then it is the new one. Return that Table's Object.
	For $i = 0 To $oTables.getCount() - 1

		For $j = 0 To UBound($asTables) - 1
			If ($asTables[$j] = $oTables.getByIndex($i).Name()) Then ExitLoop
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0)) ; Sleep every x cycles.
		Next

		If ($j = UBound($asTables)) Then ; If No matches in the original table names, then set Table Object and exit loop

			$oTable = $oTables.getByIndex($i)
			ExitLoop
		EndIf
	Next

	Return (IsObj($oTable)) ? SetError($__LOW_STATUS_SUCCESS, 0, $oTable) : SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
EndFunc   ;==>_LOWriter_DocConvertTextToTable

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocCreate
; Description ...: Open a new Libre Office Writer Document or Connect to an existing blank, unsaved, writable document.
; Syntax ........: _LOWriter_DocCreate([$bForceNew = True[, $bHidden = False]])
; Parameters ....: $bForceNew		- [optional] a boolean value. Default is True. Whether to force opening a new Writer Document instead of checking for a usable blank.
;				   $bHidden			- [optional] a boolean value. Default is False. If True opens the new document invisible or changes the existing document to invisible.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $bForceNew not a Boolean.
;				   @Error 1 @Extended 2 Return 0 = $bHidden not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failure Creating Object com.sun.star.ServiceManager.
;				   @Error 2 @Extended 2 Return 0 = Failure Creating Object com.sun.star.frame.Desktop.
;				   @Error 2 @Extended 3 Return 0 = Failure Enumerating available documents.
;				   @Error 2 @Extended 4 Return 0 = Failure Creating New Document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Document Object is still returned. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bHidden
;				   --Success--
;				   @Error 0 @Extended 1 Return Object = Successfully connected to an existing Document. Returning Document's Object
;				   @Error 0 @Extended 2 Return Object = Successfully created a new document. Returning Document's Object
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

	Local Const $iURLFrameCreate = 8 ;frame will be created if not found
	Local $aArgs[1]
	Local $iError = 0
	Local $oServiceManager, $oDesktop, $oDoc, $oEnumDoc

	If Not IsBool($bForceNew) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bHidden) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$aArgs[0] = __LOWriter_SetPropertyValue("Hidden", $bHidden)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	; If not force new, and L.O pages exist then see if any blank writer documents to use.
	If Not $bForceNew And $oDesktop.getComponents.hasElements() Then
		$oEnumDoc = $oDesktop.getComponents.createEnumeration()
		If Not IsObj($oEnumDoc) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)
		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()
			If $oDoc.supportsService("com.sun.star.text.TextDocument") _
					And Not ($oDoc.hasLocation() And $oDoc.isReadOnly()) And ($oDoc.WordCount() = 0) Then
				$oDoc.CurrentController.Frame.ContainerWindow.Visible = ($bHidden) ? False : True ; opposite value of $bHidden.
				$iError = ($oDoc.CurrentController.Frame.isHidden() = $bHidden) ? $iError : BitOR($iError, 1)
				Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, $oDoc) : SetError($__LOW_STATUS_SUCCESS, 1, $oDoc)
			EndIf
		WEnd
	EndIf

	If Not IsObj($aArgs[0]) Then Return $iError = BitOR($iError, 1)
	$oDoc = $oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", $iURLFrameCreate, $aArgs)
	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INIT_ERROR, 4, 0)
	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, $oDoc) : SetError($__LOW_STATUS_SUCCESS, 2, $oDoc)

EndFunc   ;==>_LOWriter_DocCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocCreateTextCursor
; Description ...: Create a TextCursor Object for future Textcursor related functional use.
; Syntax ........: _LOWriter_DocCreateTextCursor(Byref $oDoc[, $bCreateAtEnd = True[, $bCreateAtViewCursor = False]])
; Parameters ....: $oDoc				- [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
;                  $bCreateAtEnd		- [optional] a boolean value. Default is True. If true,  creates the new cursor at the end of the Document. Else cursor is created at the beginning.
;                  $bCreateAtViewCursor - [optional] a boolean value. Default is False. Create the Text cursor at the document's View Cursor. See Remarks
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bCreateAtEnd not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bCreateAtViewCursor not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bCreateAtEnd and $bCreateAtViewCursor both set to True, set either one to False.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve current ViewCursor Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create Text Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to create Cursor Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Current ViewCursor is in unknown data type or failed detecting what data type.
;				   --Success--
;				   @Error 0 @Extended ? Return Object = Success, Cursor object was returned. @Extended can be on of the constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3 indicating the current created cursor is in that type of data.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The cursor Created by this function in a text document, is used for inserting text, reading text, etc.
;					If you set $bCreateAtEnd to False, the new cursor is created at the beginning of the document, True creates
;					the cursor at the very end of the document. Setting $bCreateAtViewCursor to True will create a Textcursor at
;					the current ViewCursor position.
; +
;						There are two types of cursors in Word documents. The one you see, called the "ViewCursor", and the one
;					you do not see, called the "TextCursor". A "ViewCursor" is the blinking cursor you see when you are editing
;					a Word document, there is only one per document. A "TextCursor" on the other hand, is an invisible cursor
;					used for inserting text etc., into a Writer document. You can have multiple "TextCursors" per document.
; Related .......: _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocCreateTextCursor(ByRef $oDoc, $bCreateAtEnd = True, $bCreateAtViewCursor = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCursor, $oText, $oViewCursor
	Local $iCursorType = 0

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bCreateAtEnd) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bCreateAtViewCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If ($bCreateAtEnd = True) And ($bCreateAtViewCursor = True) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	If ($bCreateAtViewCursor = True) Then
		$oViewCursor = $oDoc.CurrentController.getViewCursor()
		If Not IsObj($oViewCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
		$oText = __LOWriter_CursorGetText($oDoc, $oViewCursor)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		$iCursorType = @extended
		If Not IsObj($oText) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
		If __LOWriter_IntIsBetween($iCursorType, $LOW_CURDATA_BODY_TEXT, $LOW_CURDATA_HEADER_FOOTER) Then
			$oCursor = $oText.createTextCursorByRange($oViewCursor)
		Else
			Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ; ViewCursor in unknown data type.
		EndIf

	Else
		$oText = $oDoc.getText
		If Not IsObj($oText) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
		$oCursor = $oText.createTextCursor()
		$iCursorType = $LOW_CURDATA_BODY_TEXT

		If ($bCreateAtEnd = True) Then
			$oCursor.gotoEnd(False)
		Else
			$oCursor.gotoStart(False)
		EndIf
	EndIf

	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	Return SetError($__LOW_STATUS_SUCCESS, $iCursorType, $oCursor)
EndFunc   ;==>_LOWriter_DocCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocDescription
; Description ...: Set or Retrieve Document Description properties.
; Syntax ........: _LOWriter_DocDescription(Byref $oDoc[, $sTitle = Null[, $sSubject = Null[, $aKeywords = Null[, $sComments = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
;                  $sTitle              - [optional] a string value. Default is Null. Set the Document's "Title Property. See Remarks.
;                  $sSubject            - [optional] a string value. Default is Null. Set the Document's "Subject" Property.
;                  $aKeywords           - [optional] an array of strings. Default is Null. Set the Document's "Keywords" Property.
;				   +						Input must be a single dimension Array, which will overwrite any keywords previously set. Accepts numbers also. See Remarks.
;                  $sComments           - [optional] a string value. Default is Null. Set the Document's "Comments" Property.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sTitle not a String.
;				   @Error 1 @Extended 3 Return 0 = $sSubject not a String.
;				   @Error 1 @Extended 4 Return 0 = $asKeywords not an Array.
;				   @Error 1 @Extended 5 Return 0 = $sComments not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sTitle
;				   |								2 = Error setting $sSubject
;				   |								4 = Error setting $asKeywords
;				   |								8 = Error setting $sComments
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
;				   +								"Keywords" value will be an Array, which could be empty if no keywords are presently set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: "Title" is the Title as found in File>Properties, not the Document's Title as set when saving it.
;					"Keywords" error checking only checks to make sure the input array, and the set Array of Keywords is the same size, it does not check that each element is the same.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocDescription(ByRef $oDoc, $sTitle = Null, $sSubject = Null, $asKeywords = Null, $sComments = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocProp
	Local $iError = 0
	Local $avDescription[4]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sTitle, $sSubject, $asKeywords, $sComments) Then
		__LOWriter_ArrayFill($avDescription, $oDocProp.Title(), $oDocProp.Subject(), $oDocProp.Keywords(), $oDocProp.Description())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDescription)
	EndIf

	If ($sTitle <> Null) Then
		If Not IsString($sTitle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocProp.Title = $sTitle
		$iError = ($oDocProp.Title() = $sTitle) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sSubject <> Null) Then
		If Not IsString($sSubject) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocProp.Subject = $sSubject
		$iError = ($oDocProp.Subject() = $sSubject) ? $iError : BitOR($iError, 2)
	EndIf

	If ($asKeywords <> Null) Then
		If Not IsArray($asKeywords) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oDocProp.Keywords = $asKeywords
		$iError = (UBound($oDocProp.Keywords()) = UBound($asKeywords)) ? $iError : BitOR($iError, 4)
	EndIf

	If ($sComments <> Null) Then
		If Not IsString($sComments) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocProp.Description = $sComments
		$iError = ($oDocProp.Description() = $sComments) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocDescription

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocEnumPrinters
; Description ...: Enumerates all installed printers, or current default printer.
; Syntax ........: _LOWriter_DocEnumPrinters([$bDefaultOnly = False])
; Parameters ....: $bDefaultOnly        - [optional] a boolean value. Default is False. If True, returns only the name of the current default printer. Libre 6.3 and up only.
; Return values .: Success: An array or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $bDefaultOnly Not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failure Creating "com.sun.star.ServiceManager" Object.
;				   @Error 2 @Extended 2 Return 0 = Failure creating "com.sun.star.awt.PrinterServer" Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve Default printer name.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve Array of printer names.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.1.
;				   @Error 7 @Extended 2 Return 0 = Current Libre Office version lower than 6.3.
;				   --Success--
;				   @Error 0 @Extended 1 Return String = Returning the default printer name.
;				   @Error 0 @Extended ? Return Array = Returning an array of strings containing all installed printers. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function works for LibreOffice 4.1 and Up.
; Related .......: _LOWriter_DocEnumPrintersAlt
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocEnumPrinters($bDefaultOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oPrintServer
	Local $sDefault
	Local $asPrinters[0]

	If Not __LOWriter_VersionCheck(4.1) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsBool($bDefaultOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If @error Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oPrintServer = $oServiceManager.createInstance("com.sun.star.awt.PrinterServer")
	If Not IsObj($oPrintServer) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	If $bDefaultOnly Then
		If Not __LOWriter_VersionCheck(6.3) Then Return SetError($__LOW_STATUS_VER_ERROR, 2, 0)
		$sDefault = $oPrintServer.getDefaultPrinterName()
		If IsString($sDefault) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $sDefault)
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	$asPrinters = $oPrintServer.getPrinterNames()
	If IsArray($asPrinters) Then Return SetError($__LOW_STATUS_SUCCESS, UBound($asPrinters), $asPrinters)
	Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)

EndFunc   ;==>_LOWriter_DocEnumPrinters

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocEnumPrintersAlt
; Description ...: Alternate function; Enumerates all installed printers, or current default printer.
; Syntax ........: _LOWriter_DocEnumPrintersAlt([$sPrinterName = ""[, $bReturnDefault = False]])
; Parameters ....: $sPrinterName        - [optional] a string value. Default is "". Name of the printer to list.
;				   +						Default "" returns the list of all printers.
;				   +						$sPrinterName can be a part of a printer name like "HP*". Remember the asterisk (*).
;                  $bReturnDefault      - [optional] a boolean value. Default is False.
;				   +						If True, returns only the name of the current default printer.
; Return values .: Success: Array or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sPrinterName - Not a String type variable.
;				   @Error 1 @Extended 2 Return 0 = $bReturnDefault Not a Boolean (True/False) type variable.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failure Creating Object
;				   @Error 2 @Extended 2 Return 0 = Failure retrieving printer list Object.
;				   --Printer Related Errors--
;				   @Error 6 @Extended 1 Return 0 = No default printer found.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Returnin an array of string containing all installed printers. See remarks. Number of results returned in @Extended.
;				   @Error 0 @Extended 2 Return String = Returning the default printer name. See remarks.
; Author ........: jguinch (_PrintMgr_EnumPrinter)
; Modified ......: donnyh13 - Added input error checking. Added a return default printer only option.
; Remarks .......: When $bReturnDefault is False, The function returns all installed printers for the user running the script in an array.
;					@Extended is set to the number of results. If $sPrinterName is set, the name must be exact, or no results will be found.
;					Use an asterisk (*) for partial name searches, either prefixed (*Canon), suffixed (Canon*), or both (*Canon*).
;					When $bReturnDefault is True, The function returns only the default printer's name or sets an error if no default printer is found.
; Related .......: _LOWriter_DocEnumPrinters
; Link ..........: https://www.autoitscript.com/forum/topic/155485-printers-management-udf/
; UDF title......: Printmgr.au3
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocEnumPrintersAlt($sPrinterName = "", $bReturnDefault = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asPrinterNames[10]
	Local $sFilter
	Local $iCount = 0
	Local Const $wbemFlagReturnImmediately = 0x10, $wbemFlagForwardOnly = 0x20
	Local $oWMIService, $oPrinters

	If Not IsString($sPrinterName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnDefault) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If $sPrinterName <> "" Then $sFilter = StringReplace(" Where Name like '" & StringReplace($sPrinterName, "\", "\\") & "'", "*", "%")
	$oWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	If Not IsObj($oWMIService) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oPrinters = $oWMIService.ExecQuery("Select * from Win32_Printer" & $sFilter, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If Not IsObj($oPrinters) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	For $oPrinter In $oPrinters
		Switch $bReturnDefault
			Case False
				If $iCount >= (UBound($asPrinterNames) - 1) Then ReDim $asPrinterNames[UBound($asPrinterNames) * 2]
				$asPrinterNames[$iCount] = $oPrinter.Name
				$iCount += 1

			Case True
				If $oPrinter.Default Then Return SetError($__LOW_STATUS_SUCCESS, 2, $oPrinter.Name)
		EndSwitch
	Next
	If $bReturnDefault Then Return SetError($__LOW_STATUS_PRINTER_RELATED_ERROR, 1, 0)
	ReDim $asPrinterNames[$iCount]
	Return SetError($__LOW_STATUS_SUCCESS, $iCount, $asPrinterNames)
EndFunc   ;==>_LOWriter_DocEnumPrintersAlt

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocExecuteDispatch
; Description ...: Executes a command for a document.
; Syntax ........: _LOWriter_DocExecuteDispatch(Byref $oDoc, $sDispatch)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sDispatch           - a string value. The Dispatch command to execute. See List of commands below.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sDispatch not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Succesfully executed dispatch command.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Dispatch is essentialy a simulation of the user performing an action, such as pressing Ctrl+A to select
;						all, etc.
; Dispatch Commands: 	uno:FullScreen -- Toggles full screen mode.
;						uno:ChangeCaseToLower -- Changes all selected text to lower case.  Text must be selected with the ViewCursor.
;						uno:ChangeCaseToUpper -- Changes all selected text to upper case.  Text must be selected with the ViewCursor.
;						uno:ChangeCaseRotateCase -- Cycles the Case (Title Case, Sentence case, UPPERCASE, lowercase). Text must be selected with the ViewCursor.
;						uno:ChangeCaseToSentenceCase -- Changes the sentence to Sentence case where the Viewcursor is currently positioned or has selected.
;						uno:ChangeCaseToTitleCase -- Changes the selected text to Title case. Text must be selected with the ViewCursor.
;						uno:ChangeCaseToToggleCase -- Toggles the selected text's case (A becomes a, b becomes B, etc.).Text must be selected with the ViewCursor.
;						uno:UpdateAll -- Causes all non fixed Fields, Links, Indexes, Charts etc., to be updated.
;						uno:UpdateFields -- Causes all Fields to be updated.
;						uno:UpdateAllIndexes -- Causes all Indexes to be updated.
;						uno:UpdateAllLinks -- Causes all Links to be updated.
;						uno:UpdateCharts -- Causes all Charts to be updated.
;						uno:Repaginate -- Update Page Formatting.
;						uno:ResetAttributes -- Removes all direct formatting from the selected text. Text must be selected with the ViewCursor.
;					 	uno:SwBackspace -- Simulates pressing the Backspace key.
;						uno:Delete -- Simulates pressing the Delete key.
;						uno:Paste -- Pastes the data out of the clipboard. Simulating Ctrl+V.
;						uno:PasteUnformatted -- Pastes the data out of the clipboard unformatted.
;						uno:PasteSpecial -- Simulates pasting with Ctrl+Shift+V, opens a dialog for selecting paste format.
;						uno:Copy -- Simulates Ctrl+C, copies selected data to the clipboard. Text must be selected with the ViewCursor.
;						uno:Cut -- Simulates Ctrl+X, cuts selected data, placing it into the clipboard. Text must be selected with the ViewCursor.
;						uno:SelectAll -- Simulates Ctrl+A being pressed at the ViewCursor location.
;						uno:Zoom50Percent -- Set the zoom level to 50%.
;						uno:Zoom75Percent -- Set the zoom level to 75%.
;						uno:Zoom100Percent -- Set the zoom level to 100%.
;						uno:Zoom150Percent -- Set the zoom level to 150%.
;						uno:Zoom200Percent -- Set the zoom level to 200%.
;						uno:ZoomMinus -- Decreases the zoom value to the next increment down.
;						uno:ZoomPlus -- Increases the zoom value to the next increment up.
;						uno:ZoomPageWidth -- Set zoom to fit page width.
;						uno:ZoomPage -- Set zoom to fit page.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocExecuteDispatch(ByRef $oDoc, $sDispatch)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aArray[0]
	Local $oServiceManager, $oDispatcher

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sDispatch) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$oDispatcher.executeDispatch($oDoc.CurrentController(), "." & $sDispatch, "", 0, $aArray)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)

EndFunc   ;==>_LOWriter_DocExecuteDispatch

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocExport
; Description ...: Export a Document with the specified file name to the path specified, with any parameters used.
; Syntax ........: _LOWriter_DocExport(Byref $oDoc, $sFilePath[, $bSamePath = False[, $sFilterName = ""[, $bOverwrite = Null[, $sPassword = Null]]]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFilePath      - a string value. Full path to save the document to, including Filename and extension. See Remarks.
;                  $bSamePath      - [optional] a boolean value. Default is False. Uses the path of the current document to export to. See Remarks
;                  $sFilterName    - [optional] a string value. Default is "". Filter name. "" (blank string).
;				   +					Filter is chosen automatically based on the file extension. If no extension is present, with filtername of "writer8" or if not matched to the list of extensions in this UDF, the .odt extension is used instead,
;                  $bOverwrite     - [optional] a boolean value. Default is Null. If True, file will be overwritten.
;                  $sPassword      - [optional] a string value. Default is Null. Password String to set for the document. (Not all file formats can have a Password set). "" (blank string) or Null = No Password.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFilePath Not a String.
;				   @Error 1 @Extended 3 Return 0 = $bSamePath Not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $sFilterName Not a String.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite Not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sPassword Not a String.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error Converting Path to/from L.O. URL
;				   @Error 3 @Extended 2 Return 0 = Error retrieving FilterName.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Error setting FilterName Property
;				   @Error 4 @Extended 2 Return 0 = Error setting Overwrite Property
;				   @Error 4 @Extended 3 Return 0 = Error setting Password Property
;				   --Document Errors--
;				   @Error 5 @Extended 1 Return 0 = Document has no save path, and $bSamePath is set to True.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returning save path for exported document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Does not alter the original save path (if there was one), saves a copy of the document to the new path, in the new file format if one is chosen.
;					If $bSamePath is set to True, the same save path as the current document is used.
;					You must still fill in "sFilePath" with the desired File Name and new extension, but you do not need to enter the file path.
; Related .......: _LOWriter_DocSave, _LOWriter_DocSaveAs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocExport(ByRef $oDoc, $sFilePath, $bSamePath = False, $sFilterName = "", $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aProperties[3]
	Local $sOriginalPath, $sSavePath

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bSamePath) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sFilterName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	If $bSamePath Then
		If $oDoc.hasLocation() Then
			$sOriginalPath = $oDoc.getURL()
			$sOriginalPath = StringLeft($sOriginalPath, StringInStr($sOriginalPath, "/", 0, -1)) ; Cut the original name off.
			If StringInStr($sFilePath, "\") Then $sFilePath = _LOWriter_PathConvert($sFilePath, $LOW_PATHCONV_OFFICE_RETURN) ; Convert to L.O. URL
			If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
			$sFilePath = $sOriginalPath & $sFilePath ; combine the path with the new name.
		Else
			Return SetError($__LOW_STATUS_DOC_ERROR, 1, 0)
		EndIf
	EndIf

	If Not $bSamePath Then $sFilePath = _LOWriter_PathConvert($sFilePath, $LOW_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFilterName = "") Or ($sFilterName = " ") Then $sFilterName = __LOWriter_FilterNameGet($sFilePath, True)
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)

	$aProperties[0] = __LOWriter_SetPropertyValue("FilterName", $sFilterName)
	If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOWriter_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 2, 0)
	EndIf

	If ($sPassword <> Null) Then
		If Not IsString($sPassword) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOWriter_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 3, 0)
	EndIf

	$oDoc.storeToURL($sFilePath, $aProperties)

	$sSavePath = _LOWriter_PathConvert($sFilePath, $LOW_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOWriter_DocExport

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocFindAll
; Description ...: Find all matches contained in a document of a Specified Search String.
; Syntax ........: _LOWriter_DocFindAll(Byref $oDoc, Byref $oSrchDescript, $sSearchString, Byref $atFindFormat)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $atFindFormat        - [in/out] an array of structs. An Array of formatting properties created from _LOWriter_FindFormat* functions to search for, call with an empty array to skip. Array will not be modified.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSrchDescriptObject not a Search Descriptor Object.
;				   @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;				   @Error 1 @Extended 5 Return 0 = $atFindFormat not an Array.
;				   @Error 1 @Extended 6 Return 0 = $atFindFormat does not contain an Object in the first Element.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Search did not return an Object, something went wrong.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Search was Successful, but yielded no results.
;				   @Error 0 @Extended ? Return Array = Success. Search was Successful, returning 1 dimensional array containing the objects to each match, @Exteneded is set to the number of matches.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Objects returned can be used in any of the functions accepting a Paragraph or Cursor Object etc.,
;						to modify their properties or even the text itself.
; Related .......: _LOWriter_SearchDescriptorCreate, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext,
;					_LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange, _LOWriter_FindFormatModifyAlignment,
;					_LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation,
;					_LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak,
;					_LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace,
;					_LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout,
;					_LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocFindAll(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, ByRef $atFindFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResults
	Local $aoResults[0]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsArray($atFindFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	If (UBound($atFindFormat) > 0) And Not IsObj($atFindFormat[0]) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$oSrchDescript.setSearchAttributes($atFindFormat)
	$oSrchDescript.SearchString = $sSearchString

	$oResults = $oDoc.findAll($oSrchDescript)
	If Not IsObj($oResults) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oResults.getCount() > 0) Then
		ReDim $aoResults[$oResults.getCount]
		For $i = 0 To $oResults.getCount() - 1
			$aoResults[$i] = $oResults.getByIndex($i)
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return (UBound($aoResults) > 0) ? SetError($__LOW_STATUS_SUCCESS, UBound($aoResults), $aoResults) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocFindAll

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocFindAllInRange
; Description ...: Find all occurences of a Search String in a Document in a specific selection.
; Syntax ........: _LOWriter_DocFindAllInRange(Byref $oDoc, Byref $oSrchDescript, $sSearchString, Byref $atFindFormat, Byref $oRange)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $atFindFormat        - [in/out] an array of structs. Set to an empty array to skip. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search". Array will not be modified.
;                  $oRange              - [in/out] an object. A Range, such as a cursor with Data selected, to perform the search within.
; Return values .: Success: 1 or Array..
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;				   @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;				   @Error 1 @Extended 5 Return 0 = $atFindFormat not an Array.
;				   @Error 1 @Extended 6 Return 0 = First element in $atFindFormat not an Object.
;				   @Error 1 @Extended 7 Return 0 = $oRange not set to Null and not an Object.
;				   @Error 1 @Extended 8 Return 0 = $oRange has no data selected.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Search did not return an Object, something went wrong.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Search was successful but found no matches.
;				   @Error 0 @Extended ? Return Array = Success. Search was Successful, returning 1 dimensional array containing the objects to each match, @Extended is set to the number of matches.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_SearchDescriptorCreate,
;					_LOWriter_DocFindAll, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll,
;					_LOWriter_DocReplaceAllInRange, _LOWriter_FindFormatModifyAlignment,
;					_LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation,
;					_LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak,
;					_LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace,
;					_LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout,
;					_LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocFindAllInRange(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, ByRef $atFindFormat, ByRef $oRange)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResults, $oResult, $oRangeRegion, $oResultRegion, $oText, $oRangeAnchor
	Local $aoResults[0]
	Local $iCount = 0

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsArray($atFindFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	If (UBound($atFindFormat) > 0) And Not IsObj($atFindFormat[0]) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	$oSrchDescript.setSearchAttributes($atFindFormat)

	If Not IsObj($oRange) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
	If ($oRange.IsCollapsed()) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)

	$oSrchDescript.SearchString = $sSearchString

	If $oRange.Text.supportsService("com.sun.star.text.TextFrame") Then
		$oRangeAnchor = $oRange.TextFrame.getAnchor() ; If Range is in a TextFrame, convert its position to a range in the document
	ElseIf $oRange.Text.supportsService("com.sun.star.text.Footnote") Or $oRange.Text.supportsService("com.sun.star.text.Endnote") Then
		$oRangeAnchor = $oRange.Text.Anchor()
	EndIf

	$oResults = $oDoc.findAll($oSrchDescript)
	If Not IsObj($oResults) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

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
					$oResultRegion = $oResult ;If They are, then compare the regions of that text frame.
					$oRangeRegion = $oRangeAnchor
					$oText = $oRange.Text() ; Must use the corresponding Text Object for that TextFrame as Region Compare can only compare regions contained in the same Text Object region.
				EndIf
			ElseIf $oResult.Text.supportsService("com.sun.star.text.Footnote") Or $oResult.Text.supportsService("com.sun.star.text.Endnote") And _
					$oRange.Text.supportsService("com.sun.star.text.Footnote") Or $oRange.Text.supportsService("com.sun.star.text.Endnote") Then
				If ($oDoc.Text.compareRegionEnds($oResultRegion, $oRangeAnchor) = 0) Then ;  If both Range and Result are in a Text Frame, test if they are in the same one.
					$oResultRegion = $oResult ;If They are, then compare the regions of that text frame.
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

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
		ReDim $aoResults[$iCount]
	EndIf

	Return (UBound($aoResults) > 0) ? SetError($__LOW_STATUS_SUCCESS, UBound($aoResults), $aoResults) : SetError($__LOW_STATUS_SUCCESS, 0, 1)

EndFunc   ;==>_LOWriter_DocFindAllInRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocFindNext
; Description ...: Find a Search String in a Document once or one at a time.
; Syntax ........: _LOWriter_DocFindNext(Byref $oDoc, Byref $oSrchDescript, $sSearchString, Byref $atFindFormat[, $oRange = Null[, $oLastFind = Null[, $bExhaustive = False]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $atFindFormat        - [in/out] an array of structs. Set to an empty array to skip. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search". Array will not be modified.
;                  $oRange              - [optional] an object. Default is Null. A Range, such as a cursor with Data selected, to perform the search within.
;				   +						If Null, the entire document is searched.
;                  $oLastFind           - [optional] an object. Default is Null. The last returned Object by a previous call to this function to begin the search from,
;				   +						if set to Null, the search begins at the start of the Document or selection, depending on if a Range is provided.
;                  $bExhaustive         - [optional] a boolean value. Default is False.
;				   +						If True, tests whether every result found in a document is contained in the selection or not. See remarks.
; Return values .: Success: Object or 1.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;				   @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;				   @Error 1 @Extended 5 Return 0 = $atFindFormat not an Array.
;				   @Error 1 @Extended 6 Return 0 = First element in $atFindFormat not an Object.
;				   @Error 1 @Extended 7 Return 0 = $oRange not set to Null and not an Object.
;				   @Error 1 @Extended 8 Return 0 = $oRange has no data selected.
;				   @Error 1 @Extended 9 Return 0 = $oLastFind not an Object and not set to Null, or failed to retrieve starting position from $oRange.
;				   @Error 1 @Extended 10 Return 0 = $oLastFind incorrect Object type.
;				   @Error 1 @Extended 11 Return 0 = $bExhaustive not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Search was successful but found no matches.
;				   @Error 0 @Extended 1 Return Object = Success. Search was successful, returning the resulting Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When a search is performed inside of a selection, the search may miss any footnotes/ Endnotes/ Frames
;					contained in that selection as the text of these are counted as being located at the very end/beginning of
;					a Document, thus if you are searching in the center of a document, the search will begin in the center,
;					reach the end of the selection, and stop, never reaching the foot/EndNotes etc. If $bExhaustive is set to
;					True, the search continues until the whole document has been searched, but be warned, if the search has many
;					hits, this could slow the search considerably. There is no use setting this to True in a full document
;					search.
; Related .......: _LOWriter_SearchDescriptorCreate, _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange,
;					_LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange, _LOWriter_FindFormatModifyAlignment,
;					_LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation,
;					_LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak,
;					_LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace,
;					_LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout,
;					_LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocFindNext(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, ByRef $atFindFormat, $oRange = Null, $oLastFind = Null, $bExhaustive = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResult, $oRangeRegion, $oResultRegion, $oText, $oFindRange

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsArray($atFindFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	If (UBound($atFindFormat) > 0) And Not IsObj($atFindFormat[0]) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	$oSrchDescript.setSearchAttributes($atFindFormat)

	If ($oRange <> Null) And Not IsObj($oRange) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)

	If ($oRange = Null) Then
		$oRange = $oDoc.getText.createTextCursor()
		$oRange.gotoStart(False)
		$oRange.gotoEnd(True)
	EndIf

	If ($oRange.IsCollapsed()) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)

	If ($oLastFind = Null) Then ; If Last find is not set, then set FindRange to Range beginnign or end, depending on SearchBackwards value.
		$oFindRange = ($oSrchDescript.SearchBackwards() = False) ? $oRange.Start() : $oRange.End()
	Else ;If Last find is set, set search start for beginning or end of last result, depending SearchBackwards value.
		If Not IsObj($oLastFind) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		If Not ($oLastFind.supportsService("com.sun.star.text.TextCursor")) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
		; If Search Backwards is False, then retrieve the end of the last result's range, else get the Start.
		$oFindRange = ($oSrchDescript.SearchBackwards() = False) ? $oLastFind.End() : $oLastFind.Start()
	EndIf

	If Not IsBool($bExhaustive) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 11, 0)

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
				If ($oDoc.Text.compareRegionEnds($oResultRegion, $oRangeRegion) = 0) Then ;  If both Range and Result are in a Text Frame, test if they are in the same one.
					$oResultRegion = $oResult ;If They are, then compare the regions of that text frame.
					$oRangeRegion = $oRange
					$oText = $oRange.Text() ; Must use the corresponding Text Object for that TextFrame as Region Compare can only compare regions contained in the same Text Object region.
				EndIf
			ElseIf $oResult.Text.supportsService("com.sun.star.text.Footnote") Or $oResult.Text.supportsService("com.sun.star.text.Endnote") And _
					$oRange.Text.supportsService("com.sun.star.text.Footnote") Or $oRange.Text.supportsService("com.sun.star.text.Endnote") Then
				If ($oDoc.Text.compareRegionEnds($oResultRegion, $oRangeRegion) = 0) Then ;  If both Range and Result are in a Text Frame, test if they are in the same one.
					$oResultRegion = $oResult ;If They are, then compare the regions of that text frame.
					$oRangeRegion = $oRange
					$oText = $oRange.Text() ; Must use the corresponding Text Object for that Foot/Endnote as Region Compare can only compare regions contained in the same Text Object region.
				EndIf
			EndIf

			If ($oText.compareRegionEnds($oResultRegion, $oRangeRegion) = -1) Then ; If Compare = -1, result is past range.
				If ($bExhaustive = False) Then
					$oResult = Null ;If Result is past the selection set Result to Null, but only if not doing an exhaustive search.
					ExitLoop
				Else ;If $bExhaustive is True, then update the find range.
					$oFindRange = $oResult.End()
				EndIf

			Else ;If Result is within range, exit While loop.
				ExitLoop
			EndIf
		EndIf

		$oResult = $oDoc.findNext($oFindRange, $oSrchDescript)

	WEnd

	Return (IsObj($oResult)) ? SetError($__LOW_STATUS_SUCCESS, 1, $oResult) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocFindNext

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocFooterGetTextCursor
; Description ...: Create a Text cursor in a Page Style footer for text related functions.
; Syntax ........: _LOWriter_DocFooterGetTextCursor(Byref $oPageStyle[, $bFooter = False[, $bFirstPage = False[, $bLeftPage = False[, $bRightPage = False]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or  _LOWriter_PageStyleGetObj function.
;                  $bFooter             - [optional] a boolean value. Default is False. If True, creates a text cursor for the page Footer. See Remarks.
;                  $bFirstPage          - [optional] a boolean value. Default is False. If True, creates a text cursor for the First page Footer. See Remarks.
;                  $bLeftPage           - [optional] a boolean value. Default is False. If True, creates a text cursor for Left page Footers. See Remarks.
;                  $bRightPage          - [optional] a boolean value. Default is False. If True, creates a text cursor for right page Footers. See Remarks.
; Return values .: Success: Object or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bFooter not a Boolean value.
;				   @Error 1 @Extended 3 Return 0 = $bFirstPage not a Boolean value.
;				   @Error 1 @Extended 4 Return 0 = $bLeftPage not a Boolean value.
;				   @Error 1 @Extended 5 Return 0 = $bRightPage not a Boolean value.
;				   @Error 1 @Extended 6 Return 0 = No parameters set to True.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object or Array = Success. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If more than one parameter is set to true, an array is returned with the requested objects in the order that the True parameters are listed.
;					Else the requested object is returned.
;					Note: If same content on left and right and first pages is active for the requested page style, you only need to use the $bFooter parameter,
;					the others are only for when same content on first page or same content on left and right pages is deactivated.
; Related .......: _LOWriter_PageStyleGetObj, _LOWriter_PageStyleCreate, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocFooterGetTextCursor(ByRef $oPageStyle, $bFooter = False, $bFirstPage = False, $bLeftPage = False, $bRightPage = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aoReturn[1]
	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bFooter) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bFirstPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bLeftPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bRightPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($bFooter = False) And ($bFirstPage = False) And ($bLeftPage = False) And ($bRightPage = False) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

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

	$vReturn = (UBound($aoReturn) = 1) ? $aoReturn[0] : $aoReturn ; If Array contains only one element, return it only outside of the array.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $vReturn)
EndFunc   ;==>_LOWriter_DocFooterGetTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenProp
; Description ...: Set, Retrieve, or reset a Document's General Properties.
; Syntax ........: _LOWriter_DocGenProp(Byref $oDoc[, $sNewAuthor = Null[, $iRevisions = Null[, $iEditDuration = Null[, $bApplyUserData = Null[, $bResetUserData = False]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sNewAuthor          - [optional] a string value. Default is Null. The new author of the document, can be set separately, but must be set to a string if $bResetUserData is set to True.
;                  $iRevisions          - [optional] an integer value. Default is Null. How often the document was edited and saved.
;                  $iEditDuration       - [optional] an integer value. Default is Null. The net time of editing the document (in seconds).
;                  $bApplyUserData      - [optional] a boolean value. Default is Null. If True, the user-specific settings saved within a document will be loaded with the document.
;                  $bResetUserData      - [optional] a boolean value. Default is False. Clears the document properties, such that it appears the document has just been created.
;				   +						Resets several attributes at once, as follows:
;				   +						Author is set to $sNewAuthor parameter, ($sNewAuthor MUST be setto a string).
;				   +						CreationDate is set to the current date and time; ModifiedBy is cleared;
;				   +						ModificationDate is cleared; PrintedBy is cleared; PrintDate is cleared;
;				   +						EditingDuration is cleared; EditingCycles is set to 1.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sNewAuthor not a String and $bResetUserData set to True.
;				   @Error 1 @Extended 3 Return 0 = $sNewAuthor not a String.
;				   @Error 1 @Extended 4 Return 0 = $iRevisions not an integer..
;				   @Error 1 @Extended 5 Return 0 = $iEditDuration not an integer.
;				   @Error 1 @Extended 5 Return 0 = $bApplyUserData not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Document Settings Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting  $sNewAuthor
;				   |								2 = Error setting $iRevisions
;				   |								4 = Error setting $iEditDuration
;				   |								8 = Error setting $bApplyUserData
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 0 Return 2 = Success. Document Properties were successfully Reset.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters, except $bResetUserData, as it is not a setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$oSettings = $oDoc.createInstance("com.sun.star.text.DocumentSettings")
	If Not IsObj($oSettings) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	If ($bResetUserData = True) Then
		If Not IsString($sNewAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocProp.resetUserData($sNewAuthor)
		Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	EndIf

	If __LOWriter_VarsAreNull($sNewAuthor, $iRevisions, $iEditDuration, $bApplyUserData) Then
		__LOWriter_ArrayFill($avGenProp, $oDocProp.Author(), $oDocProp.EditingCycles(), $oDocProp.EditingDuration(), $oSettings.getPropertyValue("ApplyUserData"))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avGenProp)
	EndIf

	If ($sNewAuthor <> Null) Then
		If Not IsString($sNewAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocProp.Author = $sNewAuthor
		$iError = ($oDocProp.Author() = $sNewAuthor) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iRevisions <> Null) Then
		If Not IsInt($iRevisions) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oDocProp.EditingCycles = $iRevisions
		$iError = ($oDocProp.EditingCycles() = $iRevisions) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iEditDuration <> Null) Then
		If Not IsInt($iEditDuration) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oDocProp.EditingDuration = $iEditDuration
		$iError = ($oDocProp.EditingDuration() = $iEditDuration) ? $iError : BitOR($iError, 4)
	EndIf

	If ($bApplyUserData <> Null) Then
		If Not IsBool($bApplyUserData) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oSettings.setPropertyValue("ApplyUserData", $bApplyUserData)
		$iError = ($oSettings.getPropertyValue("ApplyUserData") = $bApplyUserData) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocGenProp

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenPropCreation
; Description ...: Set or Retrieve a Document's General Creation Properties.
; Syntax ........: _LOWriter_DocGenPropCreation(Byref $oDoc[, $sAuthor = Null[, $tDateStruct = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sAuthor             - [optional] a string value. Default is Null. The initial author of the document.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sAuthor not a String.
;				   @Error 1 @Extended 3 Return 0 = $tDateStruct not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sAuthor
;				   |								2 = Error setting $tDateStruct
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sAuthor, $tDateStruct) Then
		__LOWriter_ArrayFill($avCreate, $oDocProp.Author(), $oDocProp.CreationDate())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avCreate)
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocProp.Author = $sAuthor
		$iError = ($oDocProp.Author() = $sAuthor) ? $iError : BitOR($iError, 1)
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocProp.CreationDate = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oDocProp.CreationDate(), $tDateStruct)) ? $iError : BitOR($iError, 2)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocGenPropCreation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenPropModification
; Description ...: Set or Retrieve a Document's General Modification Properties.
; Syntax ........: _LOWriter_DocGenPropModification(Byref $oDoc[, $sModifiedBy = Null[, $tDateStruct = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sModifiedBy         - [optional] a string value. Default is Null. Set the name of the last user who modified the document.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $tDateStruct not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sModifiedBy
;				   |								2 = Error setting $tDateStruct
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sModifiedBy, $tDateStruct) Then
		__LOWriter_ArrayFill($avMod, $oDocProp.ModifiedBy(), $oDocProp.ModificationDate())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avMod)
	EndIf

	If ($sModifiedBy <> Null) Then
		If Not IsString($sModifiedBy) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocProp.ModifiedBy = $sModifiedBy
		$iError = ($oDocProp.ModifiedBy() = $sModifiedBy) ? $iError : BitOR($iError, 1)
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocProp.ModificationDate = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oDocProp.ModificationDate(), $tDateStruct)) ? $iError : BitOR($iError, 2)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocGenPropModification

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenPropPrint
; Description ...: Set or Retrieve a Document's General Printed By Properties.
; Syntax ........: _LOWriter_DocGenPropPrint(Byref $oDoc[, $sPrintedBy = Null[, $tDateStruct = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sPrintedBy          - [optional] a string value. Default is Null. The name of the person who most recently printed the document.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sPrintedBy not a String.
;				   @Error 1 @Extended 3 Return 0 = $tDateStruct not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sPrintedBy
;				   |								2 = Error setting $tDateStruct
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sPrintedBy, $tDateStruct) Then
		__LOWriter_ArrayFill($avPrint, $oDocProp.PrintedBy(), $oDocProp.PrintDate())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avPrint)
	EndIf

	If ($sPrintedBy <> Null) Then
		If Not IsString($sPrintedBy) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocProp.PrintedBy = $sPrintedBy
		$iError = ($oDocProp.PrintedBy() = $sPrintedBy) ? $iError : BitOR($iError, 1)
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oDocProp.PrintDate = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oDocProp.PrintDate(), $tDateStruct)) ? $iError : BitOR($iError, 2)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocGenPropPrint

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGenPropTemplate
; Description ...: Set or Retrieve a Document's General Template Properties.
; Syntax ........: _LOWriter_DocGenPropTemplate(Byref $oDoc[, $sTemplateName = Null[, $sTemplateURL = Null[, $tDateStruct = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sTemplateName       - [optional] a string value. Default is Null. The name of the template from which the document was created.
;				   +						The value is an empty string if the document was not created from a template or if it was detached from the template
;                  $sTemplateURL        - [optional] a string value. Default is Null. The URL of the template from which the document was created.
;				   +						The value is an empty string if the document was not created from a template or if it was detached from the template.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display, created previously by _LOWriter_DateStructCreate.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sTemplateName not a String.
;				   @Error 1 @Extended 3 Return 0 = $sTemplateURL  not a String.
;				   @Error 1 @Extended 4 Return 0 = $tDateStruct not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Document Properties Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error converting Computer path to Libre Office URL.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sTemplateName
;				   |								2 = Error setting $sTemplateURL
;				   |								4 = Error setting $tDateStruct
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oDocProp = $oDoc.DocumentProperties()
	If Not IsObj($oDocProp) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sTemplateName, $sTemplateURL, $tDateStruct) Then
		__LOWriter_ArrayFill($avTemplate, $oDocProp.TemplateName(), _LOWriter_PathConvert($oDocProp.TemplateURL(), $LOW_PATHCONV_PCPATH_RETURN), _
				$oDocProp.TemplateDate())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avTemplate)
	EndIf

	If ($sTemplateName <> Null) Then
		If Not IsString($sTemplateName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oDocProp.TemplateName = $sTemplateName
		$iError = ($oDocProp.TemplateName() = $sTemplateName) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sTemplateURL <> Null) Then
		If Not IsString($sTemplateURL) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$sTemplateURL = _LOWriter_PathConvert($sTemplateURL, $LOW_PATHCONV_OFFICE_RETURN)
		If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		$oDocProp.TemplateURL = $sTemplateURL
		$iError = ($oDocProp.TemplateURL() = $sTemplateURL) ? $iError : BitOR($iError, 2)
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oDocProp.TemplateDate = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oDocProp.TemplateDate(), $tDateStruct)) ? $iError : BitOR($iError, 4)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocGenPropTemplate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetCounts
; Description ...: Returns the various counts contained in a document, such a paragraph, word etc.
; Syntax ........: _LOWriter_DocGetCounts(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object.  A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1 dimension array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Document Statistics Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Array = Success. A 1 dimension, 0 based, 9 row Array of integers, in the order described in remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:Returns a 1 dimension array with the following counts in this order: Page count; Line Count; Paragraph Count;
;					Word Count; Character Count; NonWhiteSpace Character Count; Table Count; Image Count; Object Count.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGetCounts(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aiCounts[9], $avDocStats

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	__LOWriter_ArrayFill($aiCounts, $oDoc.CurrentController.PageCount(), $oDoc.CurrentController.LineCount(), $oDoc.ParagraphCount(), _
			$oDoc.WordCount(), $oDoc.CharacterCount())

	$avDocStats = $oDoc.DocumentProperties.DocumentStatistics()
	If Not IsArray($avDocStats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($avDocStats) - 1
		If ($avDocStats[$i].Name() = "NonWhitespaceCharacterCount") Then $aiCounts[5] = $avDocStats[$i].Value()
		If ($avDocStats[$i].Name() = "TableCount") Then $aiCounts[6] = $avDocStats[$i].Value()
		If ($avDocStats[$i].Name() = "ImageCount") Then $aiCounts[7] = $avDocStats[$i].Value()
		If ($avDocStats[$i].Name() = "ObjectCount") Then $aiCounts[8] = $avDocStats[$i].Value()
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next

	Return SetError($__LOW_STATUS_SUCCESS, 0, $aiCounts)
EndFunc   ;==>_LOWriter_DocGetCounts

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetName
; Description ...: Retrieve the document's name.
; Syntax ........: _LOWriter_DocGetName(Byref $oDoc[, $bReturnFull = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bReturnFull         - [optional] a boolean value. Default is False. If True, the full window title is returned, such as is used by Autoit window related functions.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bReturnFull not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returns the document's current Name/Title
;				   @Error 0 @Extended 1 Return String = Success. Returns the document's current Window Title, which includes the document name and usually: "-LibreOffice Writer".
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnFull) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$sName = ($bReturnFull = True) ? $oDoc.CurrentController.Frame.Title() : $oDoc.Title()

	Return ($bReturnFull = True) ? SetError($__LOW_STATUS_SUCCESS, 1, $sName) : SetError($__LOW_STATUS_SUCCESS, 0, $sName)
EndFunc   ;==>_LOWriter_DocGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetPath
; Description ...: Returns a Document's current save path.
; Syntax ........: _LOWriter_DocGetPath(Byref $oDoc[, $bReturnLibreURL = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bReturnLibreURL     - [optional] a boolean value. Default is False.
;				   +						If True, returns a path in Libre Office URL format, else false returns a regular Windows path.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bReturnLibreURL not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = Document has no save path.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error converting Libre URL to Computer path format.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returns the P.C. path to the current document's save path.
;				   @Error 0 @Extended 1 Return String = Success. Returns the Libre Office URL to the current document's save path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_PathConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGetPath(ByRef $oDoc, $bReturnLibreURL = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sPath

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnLibreURL) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oDoc.hasLocation() Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If ($bReturnLibreURL = True) Then
		$sPath = $oDoc.URL()
	Else
		$sPath = $oDoc.URL()
		$sPath = _LOWriter_PathConvert($sPath, $LOW_PATHCONV_PCPATH_RETURN)
		If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return ($bReturnLibreURL = True) ? SetError($__LOW_STATUS_SUCCESS, 1, $sPath) : SetError($__LOW_STATUS_SUCCESS, 0, $sPath)
EndFunc   ;==>_LOWriter_DocGetPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetString
; Description ...: Retrieve the string of text currently selected or contained in a paragraph object.
; Syntax ........: _LOWriter_DocGetString(Byref $oObj)
; Parameters ....: $oObj             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions with Data selected, or a Paragraph Object returned from _LOWriter_ParObjCreateList function.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oObj doesn't support Paragraph Properties service.
;				   @Error 1 @Extended 3 Return 0 = $oObj is a TableCursor. Can only use View Cursor or Text Cursor.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Cursor type.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. The selected text in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Libre Office documentation states that when used in Libre Basic, GetString is limited to 64kb's in size.
;					I do not know if the same limitation applies to any outside use of GetString (such as through Autoit).
;					Also, if there are multiple selections, the returned value will be an empty string ("").
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocGetString(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If $oObj.supportsService("com.sun.star.text.TextCursor") Or $oObj.supportsService("com.sun.star.text.TextViewCursor") Then
		Local $iCursorType = __LOWriter_Internal_CursorGetType($oObj)
		If @error > 0 Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	EndIf
	Return SetError($__LOW_STATUS_SUCCESS, 0, $oObj.getString())
EndFunc   ;==>_LOWriter_DocGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocGetViewCursor
; Description ...: Retrieve the ViewCursor Object from a Document.
; Syntax ........: _LOWriter_DocGetViewCursor(Byref $oDoc)
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve ViewCursor Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return $oViewCursor Object = Success. The Object for the Document's View Cursor is returned for use in other Cursor related functions.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oViewCursor = $oDoc.CurrentController.getViewCursor()
	Return (IsObj($oViewCursor)) ? SetError($__LOW_STATUS_SUCCESS, 0, $oViewCursor) : SetError($__LOW_STATUS_INIT_ERROR, 1, 0) ; Failed to Create ViewCursor
EndFunc   ;==>_LOWriter_DocGetViewCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHasFrameName
; Description ...: Check if a Document contains a Frame with the specified name.
; Syntax ........: _LOWriter_DocHasFrameName(Byref $oDoc, $sFrameName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFrameName          - a string value. The Frame name to search for.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFrameName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Text Frames Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Shapes Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return False = Success. Search was successful, no Frames found matching $sFrameName.
;				   @Error 0 @Extended 1 Return True = Success. Search was successful, Frame found matching $sFrameName.
;				   @Error 0 @Extended 2 Return True = Success. Search was successful, Frame found matching $sFrameName listed as a shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Some document types, such as docx list frames as Shapes instead of TextFrames, so this function searches both.
; Related .......: _LOWriter_FrameDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocHasFrameName(ByRef $oDoc, $sFrameName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFrames, $oShapes

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFrameName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oFrames = $oDoc.TextFrames()
	If Not IsObj($oFrames) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($oFrames.hasByName($sFrameName)) Then Return SetError($__LOW_STATUS_SUCCESS, 1, True)

	; If No results, then search Shapes.
	$oShapes = $oDoc.DrawPage()
	If Not IsObj($oShapes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	If $oShapes.hasElements() Then
		For $i = 0 To $oShapes.getCount() - 1
			If ($oShapes.getByIndex($i).Name() = $sFrameName) Then
				If ($oShapes.getByIndex($i).supportsService("com.sun.star.drawing.Text")) And _
						($oShapes.getByIndex($i).Text.ImplementationName() = "SwXTextFrame") And Not _
						$oShapes.getByIndex($i).getPropertySetInfo().hasPropertyByName("ActualSize") Then Return SetError($__LOW_STATUS_SUCCESS, 2, True)
			EndIf

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, False) ; No matches
EndFunc   ;==>_LOWriter_DocHasFrameName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHasImageName
; Description ...: Check if a Document contains a Image with the specified name.
; Syntax ........: _LOWriter_DocHasImageName(Byref $oDoc, $sImageName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sImageName          - a string value. The Image name to search for.
; Return values .:  Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sImageName not a String.
;				   --Success--
;				   @Error 0 @Extended 0 Return False = Success. Search was successful, no Images found matching $sImageName.
;				   @Error 0 @Extended 1 Return True = Success. Search was successful, Image found matching $sImageName.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ImageDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocHasImageName(ByRef $oDoc, $sImageName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", "__LOWriter_InternalComErrorHandler")
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sImageName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($oDoc.GraphicObjects().hasByName($sImageName)) Then Return SetError($__LOW_STATUS_SUCCESS, 1, True)

	Return SetError($__LOW_STATUS_SUCCESS, 0, False) ;No matches
EndFunc   ;==>_LOWriter_DocHasImageName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHasPath
; Description ...: Returns whether a document has been saved to a location already or not.
; Syntax ........: _LOWriter_DocHasPath(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. Returns True if the document has a save location. Else False.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDoc.hasLocation())
EndFunc   ;==>_LOWriter_DocHasPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHasTableName
; Description ...: Check if a Document contains a Table with the specified name.
; Syntax ........: _LOWriter_DocHasTableName(Byref $oDoc, $sTableName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
;                  $sTableName          - a string value. The Table name to search for.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sTableName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Text Tables Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return False = Success. Search was successful, no tables found matching $sTableName.
;				   @Error 0 @Extended 1 Return True = Success. Search was successful, table found matching $sTableName.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_TableGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocHasTableName(ByRef $oDoc, $sTableName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTables

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sTableName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$oTables = $oDoc.TextTables()
	If Not IsObj($oTables) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($oTables.hasByName($sTableName)) Then Return SetError($__LOW_STATUS_SUCCESS, 1, True)

	Return SetError($__LOW_STATUS_SUCCESS, 0, False) ; No matches
EndFunc   ;==>_LOWriter_DocHasTableName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHeaderGetTextCursor
; Description ...: Create a Text cursor in a Page Style header for text related functions.
; Syntax ........: _LOWriter_DocHeaderGetTextCursor(Byref $oPageStyle[, $bHeader = False[, $bFirstPage = False[, $bLeftPage = False[, $bRightPage = False]]]])
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOWriter_PageStyleCreate, or  _LOWriter_PageStyleGetObj function.
;                  $bHeader             - [optional] a boolean value. Default is False. If True, creates a text cursor for the page header. See Remarks.
;                  $bFirstPage          - [optional] a boolean value. Default is False. If True, creates a text cursor for the First page header. See Remarks.
;                  $bLeftPage           - [optional] a boolean value. Default is False. If True, creates a text cursor for Left page headers. See Remarks.
;                  $bRightPage          - [optional] a boolean value. Default is False. If True, creates a text cursor for right page headers. See Remarks.
; Return values .: Success: Object or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oPageStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bHeader not a Boolean value.
;				   @Error 1 @Extended 3 Return 0 = $bFirstPage not a Boolean value.
;				   @Error 1 @Extended 4 Return 0 = $bLeftPage not a Boolean value.
;				   @Error 1 @Extended 5 Return 0 = $bRightPage not a Boolean value.
;				   @Error 1 @Extended 6 Return 0 = No parameters set to True.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object or Array = Success. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If more than one parameter is set to true, an array is returned with the requested objects in the order that the True parameters are listed.
;					Else the requested object is returned.
;					Note: If same content on left and right and first pages is active for the requested page style, you only need to use the $bHeader parameter,
;					the others are only for when same content on first page or same content on left and right pages is deactivated.
; Related .......: _LOWriter_PageStyleGetObj, _LOWriter_PageStyleCreate, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocHeaderGetTextCursor(ByRef $oPageStyle, $bHeader = False, $bFirstPage = False, $bLeftPage = False, $bRightPage = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aoReturn[1]
	Local $vReturn

	If Not IsObj($oPageStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bHeader) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bFirstPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bLeftPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bRightPage) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($bHeader = False) And ($bFirstPage = False) And ($bLeftPage = False) And ($bRightPage = False) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

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

	$vReturn = (UBound($aoReturn) = 1) ? $aoReturn[0] : $aoReturn ; If Array contains only one element, return it only outside of the array.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $vReturn)
EndFunc   ;==>_LOWriter_DocHeaderGetTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocHyperlinkInsert
; Description ...: Insert a hyperlink into the specified document and a cursor location or other.
; Syntax ........: _LOWriter_DocHyperlinkInsert(Byref $oDoc, Byref $oCursor, $sLinkText, $sLinkAddress[, $bInsertAtViewCursor = False[, $bOverwrite = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions. See Remarks.
;                  $sLinkText           - a string value. Link text you want displayed (Insert the URL here too if you want the link inserted raw.)
;                  $sLinkAddress        - a string value. A URL/Link.
;                  $bInsertAtViewCursor - [optional] a boolean value. Default is False. See Remarks
;                  $bOverwrite          - [optional] a boolean value. Default is False. If true, overwrites any data selected by the $oCursor.
; Return values .: Success: 1.
;				    Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				    --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc variable not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor variable not an Object, and is not set to Default keyword.
;				   @Error 1 @Extended 3 Return 0 = $sLinkText not a String variable type
;				   @Error 1 @Extended 4 Return 0 = $sLinkAddress not a String variable type
;				   @Error 1 @Extended 5 Return 0 = $bInsertAtViewCursor not a Boolean (True/False) variable type.
;				   @Error 1 @Extended 6 Return 0 = $oCursor is set to an Object variable, and $bInsertAtViewCursor is set
;				   +		to True. Change $oCursor to Default or set $bInsertAtViewCursor to False/ do not declare it.
;				   @Error 1 @Extended 7 Return 0 = $oCursor variable is a TableCursor, and cannot be used.
;				    --Initialiazation Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create Cursor Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create Text Object.
;				    --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve Cursor type.
;				   @Error 3 @Extended 2 Return 0 = Current ViewCursor is in unknown data type or failed detecting what data type.
;				    --Success--
;				   @Error 0 @Extended 1 Return 1 = Success, hyperlink was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You may call this function with an already existent cursor object, returned from a previous function, this
;					will place the Link at the current TextCursor's position. You can also set $oCursor to Default keyword, and
;					set $bInsertAtViewCursor to True. (More on View cursors and Text cursors later.) This will insert the link
;					at the current ViewCursor position. Or you can set $oCursor to Default, and leave $bInsertAtViewCursor
;					undeclared which will insert the Link at the very end of the document.
;				    	There are two types of cursors in Word documents. The one you see, called the "ViewCursor", and the one
;					you do not see, called the "TextCursor". A "ViewCursor" is the blinking cursor you see when you are editing
;					a Word document, there is only one per document. A "TextCursor" on the other hand, is an invisible cursor
;					used for inserting text etc., into a Writer document. You can have multiple "TextCursors". If You set
;					$bInsertAtViewCursor to True, the Link will be inserted at the current ViewCursor in the document.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocHyperlinkInsert(ByRef $oDoc, ByRef $oCursor, $sLinkText, $sLinkAddress, $bInsertAtViewCursor = False, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oText, $oTextCursor
	Local $iCursorType = 0

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) And ($oCursor <> Default) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sLinkText) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sLinkAddress) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bInsertAtViewCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If IsObj($oCursor) And $bInsertAtViewCursor Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)

	If IsObj($oCursor) Or $bInsertAtViewCursor Then
		$iCursorType = (IsObj($oCursor)) ? __LOWriter_Internal_CursorGetType($oCursor) : 0
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)

		If $bInsertAtViewCursor Or ($iCursorType = $LOW_CURTYPE_VIEW_CURSOR) Then
			$oTextCursor = _LOWriter_DocCreateTextCursor($oDoc, False, True) ; create new Text cursor at ViewCursor
			If Not IsObj($oTextCursor) Or @error Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
		EndIf

		$oTextCursor = ($iCursorType = $LOW_CURTYPE_TEXT_CURSOR) ? $oCursor : $oTextCursor ; If already was a TextCursor transfer to $oTextCursor

		$oText = __LOWriter_CursorGetText($oDoc, $oTextCursor)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
		If Not IsObj($oText) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	Else
		$oText = $oDoc.getText
		If Not IsObj($oText) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
		$oTextCursor = $oText.createTextCursor()
		If Not IsObj($oTextCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
		$oTextCursor.gotoEnd(False)
	EndIf

	$oText.insertString($oTextCursor, $sLinkText, False)
	With $oTextCursor
		.goLeft(StringLen($sLinkText), False)
		.goRight(StringLen($sLinkText), True)
		.HyperLinkURL = $sLinkAddress
		.collapseToEnd()
		.goRight(1, False)
	EndWith
	Return SetError($__LOW_STATUS_SUCCESS, 1, $oTextCursor)
EndFunc   ;==>_LOWriter_DocHyperlinkInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocInsertControlChar
; Description ...: Insert a control character at the cursor position.
; Syntax ........: _LOWriter_DocInsertControlChar(Byref $oDoc, Byref $oCursor, $iConChar[, $bOverwrite = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $iConChar            - an integer value (0-5). The control character to insert. See constants, $LOW_CON_CHAR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If true, and the cursor object has text selected, it is overwritten, else the character is inserted to the left of the selection.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iConChar not an Integer, less than 0 or higher than 5. See Constants, $LOW_CON_CHAR_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $oCursor is a TableCursor. Can only use View Cursor or Text Cursor.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Cursor type.
;				   @Error 3 @Extended 2 Return 0 = Error creating Text Cursor.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Control Character was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_DocInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocInsertControlChar(ByRef $oDoc, ByRef $oCursor, $iConChar, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $oTextCursor = $oCursor

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_IntIsBetween($iConChar, 0, 5) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error > 0 Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iCursorType = $LOW_CURTYPE_VIEW_CURSOR) Then $oTextCursor = _LOWriter_DocCreateTextCursor($oDoc, False, True)

	If Not IsObj($oTextCursor) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)

	$oTextCursor.Text.insertControlCharacter($oTextCursor, $iConChar, $bOverwrite)
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocInsertControlChar

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocInsertString
; Description ...: Insert a string at a cursor position.
; Syntax ........: _LOWriter_DocInsertString(Byref $oDoc, Byref $oCursor, $sString[, $bOverwrite = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $sString             - a string value. A String to insert.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If true, and the cursor object has text selected, the selection is overwritten, else the string is inserted to the left of the selection.
;				   +						If there are multiple selections, the string is inserted to the left of the last selection, and none are overwritten.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sString not a string..
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $oCursor is a TableCursor. Can only use View Cursor or Text Cursor.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Cursor type.
;				   @Error 3 @Extended 2 Return 0 = Error creating Text Cursor.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. String was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocInsertString(ByRef $oDoc, ByRef $oCursor, $sString, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $oTextCursor = $oCursor

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sString) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error > 0 Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($iCursorType = $LOW_CURTYPE_VIEW_CURSOR) Then $oTextCursor = _LOWriter_DocCreateTextCursor($oDoc, False, True)

	If Not IsObj($oTextCursor) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)

	$oTextCursor.Text.insertString($oTextCursor, $sString, $bOverwrite)
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocInsertString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocIsActive
; Description ...: Tests if called document is the active document of other Libre windows.
; Syntax ........: _LOWriter_DocIsActive(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. Returns True if document is the currently active Libre window. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note, this does NOT test if the document is the current active window in Windows, it only tests if the
;					document is the current active document among other Libre Office documents.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocIsActive(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDoc.CurrentController.Frame.isActive())
EndFunc   ;==>_LOWriter_DocIsActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocIsModified
; Description ...: Returns whether the document has been modified since being created or since the last save.
; Syntax ........: _LOWriter_DocIsModified(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. Returns True if the document has been modified since last being saved.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDoc.isModified())
EndFunc   ;==>_LOWriter_DocIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocIsReadOnly
; Description ...: Returns a boolean whether a document is currently set to ReadOnly.
; Syntax ........: _LOWriter_DocIsReadOnly(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. Returns True is document is currently Read Only, else False.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	Return SetError($__LOW_STATUS_SUCCESS, 0, $oDoc.isReadOnly())
EndFunc   ;==>_LOWriter_DocIsReadOnly

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocMaximize
; Description ...: Maximize or restore a document.
; Syntax ........: _LOWriter_DocMaximize(Byref $oDoc[, $bMaximize = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bMaximize           - [optional] a boolean value. Default is Null. If True, document window is maximized, else if false, document is restored to its previous size and location.
;				   +						If Null, returns a Boolean indicating if document is currently maximized (True).
; Return values .: Success: 1 or Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bMaximize not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Document was successfully maximized.
;				   @Error 0 @Extended 1 Return Boolean = Success. $bMaximize set to Null, returning boolean indicating if Document is currently maximized (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocMaximize(ByRef $oDoc, $bMaximize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($bMaximize = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.IsMaximized())

	If Not IsBool($bMaximize) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.IsMaximized = $bMaximize
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocMaximize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocMinimize
; Description ...: Minimize or restore a document.
; Syntax ........: _LOWriter_DocMinimize(Byref $oDoc[, $bMinimize = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bMinimize           - [optional] a boolean value. Default is Null. If True, document window is minimized, else if false, document is restored to its previous size and location.
;				   +						If Null, returns a Boolean indicating if document is currently minimized (True).
; Return values .: Success: 1 or Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bMinimize not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Document was successfully minimized.
;				   @Error 0 @Extended 1 Return Boolean = Success. $bMinimize set to Null, returning boolean indicating if Document is currently minimized (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocMinimize(ByRef $oDoc, $bMinimize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($bMinimize = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.IsMinimized())

	If Not IsBool($bMinimize) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.IsMinimized = $bMinimize
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocMinimize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocOpen
; Description ...: Open an existing Writer Document, returning its object identifier.
; Syntax ........: _LOWriter_DocOpen($sFilePath[, $bConnectIfOpen = True[, $bHidden = False[, $bReadOnly = False[, $sPassword = ""[, $bLoadAsTemplate = False[, $sFilterName = ""]]]]]])
; Parameters ....: $sFilePath           - a string value. Full path and filename of the file to be opened.
;                  $bConnectIfOpen      - [optional] a boolean value. Default is True(Connect). Whether to connect to the requested document if it is already open.
;				   +						Note: any parameters (Hidden, template etc., will not be applied when connecting)
;                  $bHidden             - [optional] a boolean value. Default is Null. If true, opens the document invisibly.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If true, opens the document as read-only.
;                  $sPassword           - [optional] a string value. Default is Null. The password that was used to read-protect the document, if any.
;                  $bLoadAsTemplate     - [optional] a boolean value. Default is Null. If true, opens the document as a Template, i.e. an untitled copy of the specified document is made instead of modifying the original document.
;                  $sFilterName         - [optional] a string value. Default is Null. Name of a LibreOffice filter to use to load the specified document.
;				   +						LibreOffice automatically selects which to use by default.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sFilePath not string, or file not found.
;				   @Error 1 @Extended 2 Return 0 = Error converting filepath to URL path.
;				   @Error 1 @Extended 3 Return 0 = $bConnectIfOpen not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bHidden not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bReadOnly not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sPassword not a string.
;				   @Error 1 @Extended 7 Return 0 = $bLoadAsTemplate not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $sFilterName not a string.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create ServiceManager Object
;				   @Error 2 @Extended 2 Return 0 = Failed to create Desktop Object
;				   @Error 2 @Extended 3 Return 0 = Failed opening or connecting to document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bHidden
;				   |								2 = Error setting $bReadOnly
;				   |								4 = Error setting $sPassword
;				   |								8 = Error setting $bLoadAsTemplate
;				   |								16 = Error setting $sFilterName
;				   --Success--
;				   @Error 0 @Extended 1 Return Object = Successfully connected to requested Document without requested parameters. Returning Document's Object.
;				   @Error 0 @Extended 2 Return Object = Successfully opened requested Document with requested parameters. Returning Document's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocCreate, _LOWriter_DocClose, _LOWriter_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocOpen($sFilePath, $bConnectIfOpen = True, $bHidden = Null, $bReadOnly = Null, $sPassword = Null, $bLoadAsTemplate = Null, $sFilterName = Null)
	Local Const $iURLFrameCreate = 8 ;frame will be created if not found
	Local $iError = 0
	Local $oDoc, $oServiceManager, $oDesktop
	Local $aoProperties[0]
	Local $vProperty
	Local $sFileURL

	Local $oComError = ObjEvent("AutoIt.Error", "__LOWriter_InternalComErrorHandler")

	If Not IsString($sFilePath) Or Not FileExists($sFilePath) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$sFileURL = _LOWriter_PathConvert($sFilePath, $LOW_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bConnectIfOpen) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	If Not __LOWriter_VarsAreNull($bHidden, $bReadOnly, $sPassword, $bLoadAsTemplate, $sFilterName) Then

		If ($bHidden <> Null) Then
			If Not IsBool($bHidden) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			$vProperty = __LOWriter_SetPropertyValue("Hidden", $bHidden)
			If @error Then $iError = BitOR($iError, 1)
			If Not BitAND($iError, 1) Then __LOWriter_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($bReadOnly <> Null) Then
			If Not IsBool($bReadOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			$vProperty = __LOWriter_SetPropertyValue("ReadOnly", $bReadOnly)
			If @error Then $iError = BitOR($iError, 2)
			If Not BitAND($iError, 2) Then __LOWriter_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($sPassword <> Null) Then
			If Not IsString($sPassword) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			$vProperty = __LOWriter_SetPropertyValue("Password", $sPassword)
			If @error Then $iError = BitOR($iError, 4)
			If Not BitAND($iError, 4) Then __LOWriter_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($bLoadAsTemplate <> Null) Then
			If Not IsBool($bLoadAsTemplate) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
			$vProperty = __LOWriter_SetPropertyValue("AsTemplate", $bLoadAsTemplate)
			If @error Then $iError = BitOR($iError, 8)
			If Not BitAND($iError, 8) Then __LOWriter_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($sFilterName <> Null) Then
			If Not IsString($sFilterName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
			$vProperty = __LOWriter_SetPropertyValue("FilterName", $sFilterName)
			If @error Then $iError = BitOR($iError, 16)
			If Not BitAND($iError, 16) Then __LOWriter_AddTo1DArray($aoProperties, $vProperty)
		EndIf

	EndIf

	$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
	If StringInStr($oComError.description, """type detection failed""") And $bConnectIfOpen Then
		ReDim $aoProperties[0]
		$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
		If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

		Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, $oDoc) : SetError($__LOW_STATUS_SUCCESS, 1, $oDoc)
	EndIf

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)
	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, $oDoc) : SetError($__LOW_STATUS_SUCCESS, 2, $oDoc)
EndFunc   ;==>_LOWriter_DocOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPosAndSize
; Description ...: Reposition and resize a document window.
; Syntax ........: _LOWriter_DocPosAndSize(Byref $oDoc[, $iX = Null[, $iY = Null[, $iWidth = Null[, $iHeight = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iX                  - [optional] an integer value. Default is Null. The X coordinate of the window.
;                  $iY                  - [optional] an integer value. Default is Null. The Y coordinate of the window.
;                  $iWidth              - [optional] an integer value. Default is Null. The width of the window, in pixels(?).
;                  $iHeight             - [optional] an integer value. Default is Null. The height of the window, in pixels(?).
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iX not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iY not an Integer.
;				   @Error 1 @Extended 4 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iHeight not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Position and Size Structure Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Position and Size Structure Object for error checking.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting iX
;				   |								2 = Error setting $iY
;				   |								4 = Error setting $iWidth
;				   |								8 = Error setting $iHeight
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note, X & Y, on my computer at least, seem to go no lower than 8(X) and 30(Y), if you enter lower than this, it will cause a "property setting Error".
;					If you want more accurate functionality, use the "WinMove" Autoit function.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$tWindowSize = $oDoc.CurrentController.Frame.ContainerWindow.getPosSize()
	If Not IsObj($tWindowSize) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iX, $iY, $iWidth, $iHeight) Then
		__LOWriter_ArrayFill($aiWinPosSize, $tWindowSize.X(), $tWindowSize.Y(), $tWindowSize.Width(), $tWindowSize.Height())
		Return SetError($__LOW_STATUS_SUCCESS, 2, $aiWinPosSize)
	EndIf

	If ($iX <> Null) Then
		If Not IsInt($iX) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$tWindowSize.X = $iX
	EndIf

	If ($iY <> Null) Then
		If Not IsInt($iY) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$tWindowSize.Y = $iY
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$tWindowSize.Width = $iWidth
	EndIf

	If ($iHeight <> Null) Then
		If Not IsInt($iHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$tWindowSize.Height = $iHeight
	EndIf

	$oDoc.CurrentController.Frame.ContainerWindow.setPosSize($tWindowSize.X, $tWindowSize.Y, $tWindowSize.Width, $tWindowSize.Height, $iPosSize)

	$tWindowSize = $oDoc.CurrentController.Frame.ContainerWindow.getPosSize()
	If Not IsObj($tWindowSize) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$iError = ($iX = Null) ? $iError : ($tWindowSize.X() = $iX) ? $iError : BitOR($iError, 1)
	$iError = ($iY = Null) ? $iError : ($tWindowSize.Y() = $iY) ? $iError : BitOR($iError, 2)
	$iError = ($iWidth = Null) ? $iError : ($tWindowSize.Width() = $iWidth) ? $iError : BitOR($iError, 4)
	$iError = ($iHeight = Null) ? $iError : ($tWindowSize.Height() = $iHeight) ? $iError : BitOR($iError, 8)

	Return ($iError = 0) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0)
EndFunc   ;==>_LOWriter_DocPosAndSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrint
; Description ...: Print a document using the specified settings.
; Syntax ........: _LOWriter_DocPrint(Byref $oDoc[, $iCopies = 1[, $bCollate = True[, $vPages = "ALL"[, $bWait = True[, $iDuplexMode = $LOW_DUPLEX_OFF[, $sPrinter = ""[, $sFilePathName = ""]]]]]]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iCopies             - [optional] an integer value. Default is 1. Specifies the number of copies to print.
;                  $bCollate            - [optional] a boolean value. Default is True. Advises the printer to collate the pages of the copies.
;                  $vPages              - [optional] a String or Integer value. Default is "ALL". Specifies which pages to print. See remarks.
;                  $bWait               - [optional] a boolean value. Default is True. If True, the corresponding print request will be executed synchronous. Default is the asynchronous print mode.
;				   +						ATTENTION: Setting this field to True is highly recommended. Otherwise following actions (as e.g. closing the Document) can fail.
;                  $iDuplexMode         - [optional] an integer value (0-3). Default is $__g_iDuplexOFF. Determines the duplex mode for the print job. See Constants, $LOW_DUPLEX_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPrinter            - [optional] a string value. Default is "". Printer name. If left blank, or if printer name is not found, default printer is used.
;                  $sFilePathName       - [optional] a string value. Default is "". Specifies the name of a file to print to. Creates a .prn file at the given Path. Must include the desired path destination with file name.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iCopies not a Integer.
;				   @Error 1 @Extended 3 Return 0 = $bCollate not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $vPages Not an Integer or String.
;				   @Error 1 @Extended 5 Return 0 = $vPages contains invalid characters, a-z, or a period(.).
;				   @Error 1 @Extended 6 Return 0 = $bWait not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iDuplexMode not an Integer, less than 0 or greater than 3. See Constants, $LOW_DUPLEX_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 8 Return 0 = $sPrinter Not a String.
;				   @Error 1 @Extended 9 Return 0 = $sFilePathName Not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Error setting Printer "Name" setting.
;				   @Error 4 @Extended 2 Return 0 = Error setting "Copies" setting.
;				   @Error 4 @Extended 3 Return 0 = Error setting "Collate" setting.
;				   @Error 4 @Extended 4 Return 0 = Error setting "Wait" setting.
;				   @Error 4 @Extended 5 Return 0 = Error setting "DuplexMode" setting.
;				   @Error 4 @Extended 6 Return 0 = Error setting "Pages" setting.
;				   @Error 4 @Extended 7 Return 0 = Error converting PrintToFile Path.
;				   @Error 4 @Extended 8 Return 0 = Error setting "PrintToFile" setting.
;				   @Error 4 @Extended 3 Return 0 = Error setting "Collate" setting.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success Document was successfully printed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Based on OOoCalc UDF Print function by GMK.
;					$vPages range is given as entered in the user interface. For example: "1-4,10" to print the pages 1 to 4 and 10.
;					Default is "ALL". Must be in String format to accept more than just a single page number.
;					i.e. This will work: "1-6,12,27" This will Not 1-6,12,27. This will work: "7", This will also: 7.
; Related .......:_LOWriter_DocEnumPrintersAlt, _LOWriter_DocEnumPrinters, _LOWriter_DocPrintSizeSettings,
;					_LOWriter_DocPrintPageSettings, _LOWriter_DocPrintMiscSettings, _LOWriter_DocPrintIncludedSettings
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocPrint(ByRef $oDoc, $iCopies = 1, $bCollate = True, $vPages = "ALL", $bWait = True, $iDuplexMode = $LOW_DUPLEX_OFF, $sPrinter = "", $sFilePathName = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $STR_STRIPLEADING = 1, $STR_STRIPTRAILING = 2, $STR_STRIPALL = 8
	Local $avPrintOpt[4], $asSetPrinterOpt[1]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iCopies) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bCollate) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If Not IsInt($vPages) And Not IsString($vPages) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	$vPages = (IsString($vPages)) ? StringStripWS($vPages, $STR_STRIPALL) : $vPages
	If IsString($vPages) And Not ($vPages = "ALL") Then
		If StringRegExp($vPages, "[[:alpha:]]|[\.]") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	EndIf
	If Not IsBool($bWait) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	If Not __LOWriter_IntIsBetween($iDuplexMode, $LOW_DUPLEX_OFF, $LOW_DUPLEX_SHORT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
	If Not IsString($sPrinter) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
	$sPrinter = StringStripWS(StringStripWS($sPrinter, $STR_STRIPTRAILING), $STR_STRIPLEADING)
	If Not IsString($sFilePathName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
	$sFilePathName = StringStripWS(StringStripWS($sFilePathName, $STR_STRIPTRAILING), $STR_STRIPLEADING)
	If $sPrinter <> "" Then
		$asSetPrinterOpt[0] = __LOWriter_SetPropertyValue("Name", $sPrinter)
		If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)
		$oDoc.setPrinter($asSetPrinterOpt)
	EndIf
	$avPrintOpt[0] = __LOWriter_SetPropertyValue("CopyCount", $iCopies)
	If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 2, 0)
	$avPrintOpt[1] = __LOWriter_SetPropertyValue("Collate", $bCollate)
	If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 3, 0)
	$avPrintOpt[2] = __LOWriter_SetPropertyValue("Wait", $bWait)
	If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 4, 0)
	$avPrintOpt[3] = __LOWriter_SetPropertyValue("DuplexMode", $iDuplexMode)
	If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 5, 0)
	If $vPages <> "ALL" Then
		ReDim $avPrintOpt[UBound($avPrintOpt) + 1]
		$avPrintOpt[UBound($avPrintOpt) - 1] = __LOWriter_SetPropertyValue("Pages", $vPages)
		If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 6, 0)
	EndIf
	If $sFilePathName <> "" Then
		$sFilePathName = $sFilePathName & ".prn"
		$sFilePathName = _LOWriter_PathConvert($sFilePathName, $LOW_PATHCONV_OFFICE_RETURN)
		If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 7, 0)
		ReDim $avPrintOpt[UBound($avPrintOpt) + 1]
		$avPrintOpt[UBound($avPrintOpt) - 1] = __LOWriter_SetPropertyValue("FileName", $sFilePathName)
		If @error Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 8, 0)
	EndIf
	$oDoc.print($avPrintOpt)
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocPrint

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrintIncludedSettings
; Description ...: Set or Retrieve setting related to what is included in printing.
; Syntax ........: _LOWriter_DocPrintIncludedSettings(Byref $oDoc[, $bGraphics = Null[, $bControls = Null[, $bDrawings = Null[, $bTables = Null[, $bHiddenText = Null]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bGraphics           - [optional] a boolean value. Default is Null. If True, the graphics of the document are printed.
;                  $bControls           - [optional] a boolean value. Default is Null. If True, the form control fields of the document are printed.
;                  $bDrawings           - [optional] a boolean value. Default is Null. If True, the drawings of the document are printed.
;                  $bTables             - [optional] a boolean value. Default is Null. If True, the Tables of the document are printed.
;                  $bHiddenText         - [optional] a boolean value. Default is Null. If True, prints text that is marked as hidden.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bGraphics not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bControls not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bDrawings not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bTables not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $bHiddenText not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.DocumentSettings" Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bGraphics
;				   |								2 = Error setting $bControls
;				   |								4 = Error setting $bDrawings
;				   |								8 = Error setting $bTables
;				   |								16 = Error setting $bHiddenText
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oSettings = $oDoc.createInstance("com.sun.star.text.DocumentSettings")
	If Not IsObj($oSettings) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bGraphics, $bControls, $bDrawings, $bTables, $bHiddenText) Then
		__LOWriter_ArrayFill($abPrintSettings, $oSettings.getPropertyValue("PrintGraphics"), $oSettings.getPropertyValue("PrintControls"), _
				$oSettings.getPropertyValue("PrintDrawings"), $oSettings.getPropertyValue("PrintTables"), $oSettings.getPropertyValue("PrintHiddenText"))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $abPrintSettings)
	EndIf

	If ($bGraphics <> Null) Then
		If Not IsBool($bGraphics) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oSettings.setPropertyValue("PrintGraphics", $bGraphics)
		$iError = ($oSettings.getPropertyValue("PrintGraphics") = $bGraphics) ? $iError : BitOR($iError, 1)
	EndIf

	If ($bControls <> Null) Then
		If Not IsBool($bControls) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oSettings.setPropertyValue("PrintControls", $bControls)
		$iError = ($oSettings.getPropertyValue("PrintControls") = $bControls) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bDrawings <> Null) Then
		If Not IsBool($bDrawings) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oSettings.setPropertyValue("PrintDrawings", $bDrawings)
		$iError = ($oSettings.getPropertyValue("PrintDrawings") = $bDrawings) ? $iError : BitOR($iError, 4)
	EndIf

	If ($bTables <> Null) Then
		If Not IsBool($bTables) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oSettings.setPropertyValue("PrintTables", $bTables)
		$iError = ($oSettings.getPropertyValue("PrintTables") = $bTables) ? $iError : BitOR($iError, 8)
	EndIf

	If ($bHiddenText <> Null) Then
		If Not IsBool($bHiddenText) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oSettings.setPropertyValue("PrintHiddenText", $bHiddenText)
		$iError = ($oSettings.getPropertyValue("PrintHiddenText") = $bHiddenText) ? $iError : BitOR($iError, 16)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)

EndFunc   ;==>_LOWriter_DocPrintIncludedSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrintMiscSettings
; Description ...: Set or Retrieve Miscellaneous Printing related settings.
; Syntax ........: _LOWriter_DocPrintMiscSettings(Byref $oDoc[, $iPaperOrient = Null[, $sPrinterName = Null[, $iCommentsMode = Null[, $bBrochure = Null[, $bBrochureRTL = Null[, $bReversed = Null]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iPaperOrient        - [optional] an integer value (0-1). Default is Null. The orientation of the paper. See Constants, $LOW_PAPER_ORIENT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPrinterName        - [optional] a string value. Default is Null. The Name of the Printer to set as the printer to send print jobs to.
;                  $iCommentsMode       - [optional] an integer value (0-3). Default is Null. Where to print comments (if any). See Constants, $LOW_PRINT_NOTES_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bBrochure           - [optional] a boolean value. Default is Null. If True, prints the document in brochure format.
;                  $bBrochureRTL        - [optional] a boolean value. Default is Null. If True, prints the document in brochure Right to Left format.
;                  $bReversed           - [optional] a boolean value. Default is Null. If True, prints pages in reverse order.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iPaperOrient not an integer, less than 0 or greater than 1. See Constants, $LOW_PAPER_ORIENT_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $sPrinterName not a string.
;				   @Error 1 @Extended 4 Return 0 = $iCommentsMode not an integer, less than 0, or greater than 3. See Constants, $LOW_PRINT_NOTES_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 5 Return 0 = $bBrochure not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $bBrochureRTL not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $bReversed not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.DocumentSettings" Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving setting value of "CanSetPaperOrientation" from Printer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iPaperOrient
;				   |								2 = Error setting $sPrinterName
;				   |								4 = Error setting $iCommentsMode
;				   |								8 = Error setting $bBrochure
;				   |								16 = Error setting $bBrochureRTL
;				   |								32 = Error setting $bReversed
;				    --Printer Related Errors--
;				   @Error 6 @Extended 1 Return 0 = Printer does not allow changing paper orientation.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_DocPrintSizeSettings, _LOWriter_DocPrintPageSettings, _LOWriter_DocPrintIncludedSettings
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocPrintMiscSettings(ByRef $oDoc, $iPaperOrient = Null, $sPrinterName = Null, $iCommentsMode = Null, $bBrochure = Null, $bBrochureRTL = Null, $bReversed = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $STR_STRIPLEADING = 1, $STR_STRIPTRAILING = 2
	Local $iError = 0
	Local $oSettings
	Local $bCanSetPaperOrientation = False
	Local $aoSetting[1]
	Local $avPrintSettings[6]

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oSettings = $oDoc.createInstance("com.sun.star.text.DocumentSettings")
	If Not IsObj($oSettings) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iPaperOrient, $sPrinterName, $iCommentsMode, $bBrochure, $bReversed) Then
		__LOWriter_ArrayFill($avPrintSettings, __LOWriter_GetPrinterSetting($oDoc, "PaperOrientation"), _
				__LOWriter_GetPrinterSetting($oDoc, "Name"), $oSettings.getPropertyValue("PrintAnnotationMode"), _
				$oSettings.getPropertyValue("PrintProspect"), $oSettings.getPropertyValue("PrintProspectRTL"), _
				$oSettings.getPropertyValue("PrintReversed"))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avPrintSettings)
	EndIf

	$bCanSetPaperOrientation = __LOWriter_GetPrinterSetting($oDoc, "CanSetPaperOrientation")
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	If ($iPaperOrient <> Null) Then
		If Not __LOWriter_IntIsBetween($iPaperOrient, $LOW_PAPER_ORIENT_PORTRAIT, $LOW_PAPER_ORIENT_LANDSCAPE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		If $bCanSetPaperOrientation Then
			$aoSetting[0] = __LOWriter_SetPropertyValue("PaperOrientation", $iPaperOrient)
			$oDoc.setPrinter($aoSetting)
			$iError = (__LOWriter_GetPrinterSetting($oDoc, "PaperOrientation") = $iPaperOrient) ? $iError : BitOR($iError, 1)
		Else
			Return SetError($__LOW_STATUS_PRINTER_RELATED_ERROR, 1, 0)
		EndIf
	EndIf

	If ($sPrinterName <> Null) Then
		If Not IsString($sPrinterName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$sPrinterName = StringStripWS(StringStripWS($sPrinterName, $STR_STRIPTRAILING), $STR_STRIPLEADING)
		$aoSetting[0] = __LOWriter_SetPropertyValue("Name", $sPrinterName)
		$oDoc.setPrinter($aoSetting)
		$iError = (__LOWriter_GetPrinterSetting($oDoc, "Name") = $sPrinterName) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iCommentsMode <> Null) Then
		If Not __LOWriter_IntIsBetween($iCommentsMode, $LOW_PRINT_NOTES_NONE, $LOW_PRINT_NOTES_NEXT_PAGE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oSettings.setPropertyValue("PrintAnnotationMode", $iCommentsMode)
		$iError = ($oSettings.getPropertyValue("PrintAnnotationMode") = $iCommentsMode) ? $iError : BitOR($iError, 4)
	EndIf

	If ($bBrochure <> Null) Then
		If Not IsBool($bBrochure) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oSettings.setPropertyValue("PrintProspect", $bBrochure)
		$iError = ($oSettings.getPropertyValue("PrintProspect") = $bBrochure) ? $iError : BitOR($iError, 8)
	EndIf

	If ($bBrochureRTL <> Null) Then
		If Not IsBool($bBrochureRTL) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oSettings.setPropertyValue("PrintProspectRTL", $bBrochureRTL)
		$iError = ($oSettings.getPropertyValue("PrintProspectRTL") = $bBrochureRTL) ? $iError : BitOR($iError, 16)
	EndIf

	If ($bReversed <> Null) Then
		If Not IsBool($bReversed) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oSettings.setPropertyValue("PrintReversed", $bReversed)
		$iError = ($oSettings.getPropertyValue("PrintReversed") = $bReversed) ? $iError : BitOR($iError, 32)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocPrintMiscSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrintPageSettings
; Description ...: Set or Retrieve settings Page related print settings.
; Syntax ........: _LOWriter_DocPrintPageSettings(Byref $oDoc[, $bBlackOnly = Null[, $bLeftOnly = Null[, $bRightOnly = Null[, $bBackground = Null[, $bEmptyPages = Null]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bBlackOnly          - [optional] a boolean value. Default is Null. If True, prints all text in black only.
;                  $bLeftOnly           - [optional] a boolean value. Default is Null. If True, prints only Left(Even) pages. If both $bLeftOnly and $bRightOnly are false, both Left and Right pages are printed.
;                  $bRightOnly          - [optional] a boolean value. Default is Null. If True, prints only Right(Odd) pages. If both $bLeftOnly and $bRightOnly are false, both Left and Right pages are printed.
;                  $bBackground         - [optional] a boolean value. Default is Null. If true, prints colors and objects that are inserted to the background of the page.
;                  $bEmptyPages         - [optional] a boolean value. Default is Null. If true, automatically inserted blank pages are printed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bBlackOnly not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bLeftOnly not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bRightOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bBackground not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $bEmptyPages not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.DocumentSettings" Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bBlackOnly
;				   |								2 = Error setting $bLeftOnly
;				   |								4 = Error setting $bRightOnly
;				   |								8 = Error setting $bBackground
;				   |								16 = Error setting $bEmptyPages
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oSettings = $oDoc.createInstance("com.sun.star.text.DocumentSettings")
	If Not IsObj($oSettings) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bBlackOnly, $bLeftOnly, $bRightOnly, $bBackground, $bEmptyPages) Then
		__LOWriter_ArrayFill($abPrintSettings, $oSettings.getPropertyValue("PrintBlackFonts"), $oSettings.getPropertyValue("PrintLeftPages"), _
				$oSettings.getPropertyValue("PrintRightPages"), $oSettings.getPropertyValue("PrintPageBackground"), _
				$oSettings.getPropertyValue("PrintEmptyPages"))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $abPrintSettings)
	EndIf

	If ($bBlackOnly <> Null) Then
		If Not IsBool($bBlackOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$oSettings.setPropertyValue("PrintBlackFonts", $bBlackOnly)
		$iError = ($oSettings.getPropertyValue("PrintBlackFonts") = $bBlackOnly) ? $iError : BitOR($iError, 1)
	EndIf

	If ($bLeftOnly <> Null) Then
		If Not IsBool($bLeftOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oSettings.setPropertyValue("PrintLeftPages", $bLeftOnly)
		$iError = ($oSettings.getPropertyValue("PrintLeftPages") = $bLeftOnly) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bRightOnly <> Null) Then
		If Not IsBool($bRightOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oSettings.setPropertyValue("PrintRightPages", $bRightOnly)
		$iError = ($oSettings.getPropertyValue("PrintRightPages") = $bRightOnly) ? $iError : BitOR($iError, 4)
	EndIf

	If ($bBackground <> Null) Then
		If Not IsBool($bBackground) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oSettings.setPropertyValue("PrintPageBackground", $bBackground)
		$iError = ($oSettings.getPropertyValue("PrintPageBackground") = $bBackground) ? $iError : BitOR($iError, 8)
	EndIf

	If ($bEmptyPages <> Null) Then
		If Not IsBool($bEmptyPages) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oSettings.setPropertyValue("PrintEmptyPages", $bEmptyPages)
		$iError = ($oSettings.getPropertyValue("PrintEmptyPages") = $bEmptyPages) ? $iError : BitOR($iError, 16)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)

EndFunc   ;==>_LOWriter_DocPrintPageSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocPrintSizeSettings
; Description ...: Set or Retrieve Print Paper size settings.
; Syntax ........: _LOWriter_DocPrintSizeSettings(Byref $oDoc[, $iPaperFormat = Null[, $iPaperWidth = Null[, $iPaperHeight = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iPaperFormat        - [optional] an integer value (0-8). Default is Null. Specifies a predefined paper size or if the paper size is a user-defined size. See constants, $LOW_PAPER_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iPaperWidth         - [optional] an integer value. Default is Null. Specifies the size of the paper in micrometers. Can be a custom value or one of the constants, $LOW_PAPER_WIDTH_* as defined in LibreOfficeWriter_Constants.au3.
;				   +						Note: for some reason, setting this setting modifies the document page size also, I am unsure why.
;                  $iPaperHeight        - [optional] an integer value. Default is Null. Specifies the size of the paper in micrometers. Can be a custom value or one of the constants, $LOW_PAPER_HEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;				   +						Note: for some reason, setting this setting modifies the document page size also, I am unsure why.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iPaperFormat not an integer, less than 0 or greater than 8. See constants, $LOW_PAPER_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iPaperWidth not an integer, and not set to Null keyword.
;				   @Error 1 @Extended 4 Return 0 = $iPaperHeight not an integer, and not set to Null keyword.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.awt.Size" Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iPaperFormat
;				   |								2 = Error setting $iPaperWidth
;				   |								4 = Error setting $iPaperHeight
;				    --Printer Related Errors--
;				   @Error 6 @Extended 1 Return 0 = Printer doesn't allow paper format to be set.
;				   @Error 6 @Extended 2 Return 0 = Printer doesn't allow paper size to be set.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Due to slight inaccuracies in unit conversion, there may be false errors thrown while attempting to
;					set paper size.
;					Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DocPrintPageSettings,
;					_LOWriter_DocPrintMiscSettings, _LOWriter_DocPrintIncludedSettings
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iPaperFormat, $iPaperWidth, $iPaperHeight) Then
		__LOWriter_ArrayFill($aiPrintSettings, __LOWriter_GetPrinterSetting($oDoc, "PaperFormat"), _
				__LOWriter_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Width(), $__LOWCONST_CONVERT_TWIPS_UM), _
				__LOWriter_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Height(), $__LOWCONST_CONVERT_TWIPS_UM))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiPrintSettings)
	EndIf

	$bCanSetPaperFormat = __LOWriter_GetPrinterSetting($oDoc, "CanSetPaperFormat")
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	$bCanSetPaperSize = __LOWriter_GetPrinterSetting($oDoc, "CanSetPaperSize")
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)

	If ($iPaperFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iPaperFormat, $LOW_PAPER_A3, $LOW_PAPER_USER_DEFINED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		If $bCanSetPaperFormat Then
			$aoSetting[0] = __LOWriter_SetPropertyValue("PaperFormat", $iPaperFormat)
			$oDoc.setPrinter($aoSetting)
			$iError = (__LOWriter_GetPrinterSetting($oDoc, "PaperFormat") = $iPaperFormat) ? $iError : BitOR($iError, 1)
		Else
			Return SetError($__LOW_STATUS_PRINTER_RELATED_ERROR, 1, 0)
		EndIf
	EndIf

	If ($iPaperWidth <> Null) Or ($iPaperHeight <> Null) Then
		If $bCanSetPaperSize Then
			If Not IsInt($iPaperWidth) And ($iPaperWidth <> Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			If Not IsInt($iPaperHeight) And ($iPaperHeight <> Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

			; Set in uM but retrieved in TWIPS
			$tSize = __LOWriter_CreateStruct("com.sun.star.awt.Size")
			If Not IsObj($tSize) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
			$tSize.Width = ($iPaperWidth = Null) ? __LOWriter_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Width(), $__LOWCONST_CONVERT_TWIPS_UM) : $iPaperWidth
			$tSize.Height = ($iPaperWidth = Null) ? __LOWriter_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Height(), $__LOWCONST_CONVERT_TWIPS_UM) : $iPaperHeight
			$aoSetting[0] = __LOWriter_SetPropertyValue("PaperSize", $tSize)
			$oDoc.setPrinter($aoSetting)

			$iError = ($iPaperWidth = Null) ? $iError : (__LOWriter_IntIsBetween(__LOWriter_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Width(), $__LOWCONST_CONVERT_TWIPS_UM), $iPaperWidth - 2, $iPaperWidth + 2)) ? $iError : BitOR($iError, 2)
			$iError = ($iPaperHeight = Null) ? $iError : (__LOWriter_IntIsBetween(__LOWriter_UnitConvert(__LOWriter_GetPrinterSetting($oDoc, "PaperSize").Height(), $__LOWCONST_CONVERT_TWIPS_UM), $iPaperHeight - 2, $iPaperHeight + 2)) ? $iError : BitOR($iError, 4)

		Else
			Return SetError($__LOW_STATUS_PRINTER_RELATED_ERROR, 2, 0)
		EndIf
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)

EndFunc   ;==>_LOWriter_DocPrintSizeSettings

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocRedo
; Description ...: Perform one Redo action for a document.
; Syntax ........: _LOWriter_DocRedo(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Document does not have a redo action to perform.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Successfully performed a redo action.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocUndo, _LOWriter_DocRedoIsPossible, _LOWriter_DocRedoGetAllActionTitles,
;					_LOWriter_DocRedoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocRedo(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isRedoPossible()) Then
		$oDoc.UndoManager.Redo()
		Return SetError($__LOW_STATUS_SUCCESS, 1, 0)
	Else
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOWriter_DocRedo

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocRedoCurActionTitle
; Description ...: Retrieve the current Redo action Title.
; Syntax ........: _LOWriter_DocRedoCurActionTitle(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Document does not have a redo action available.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Returns the current available redo action Title in String format.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isRedoPossible()) Then
		Return SetError($__LOW_STATUS_SUCCESS, 1, $oDoc.UndoManager.getCurrentRedoActionTitle())
	Else
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOWriter_DocRedoCurActionTitle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocRedoGetAllActionTitles
; Description ...: Retrieve all available Redo action Titles.
; Syntax ........: _LOWriter_DocRedoGetAllActionTitles(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Document does not have any redo actions available.
;				   --Success--
;				   @Error 0 @Extended 0 Return Array = Returns all available redo action Titles in an array of Strings.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isRedoPossible()) Then
		Return SetError($__LOW_STATUS_SUCCESS, 1, $oDoc.UndoManager.getAllRedoActionTitles())
	Else
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOWriter_DocRedoGetAllActionTitles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocRedoIsPossible
; Description ...: Test whether a Redo action is available to perform for a document.
; Syntax ........: _LOWriter_DocRedoIsPossible(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If the document has a redo action to perform, True is returned, else False.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 1, $oDoc.UndoManager.isRedoPossible())
EndFunc   ;==>_LOWriter_DocRedoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocReplaceAll
; Description ...:
; Syntax ........: _LOWriter_DocReplaceAll(Byref $oDoc, Byref $oSrchDescript, $sSearchString, $sReplaceString, Byref $atFindFormat, Byref $atReplaceFormat)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - [in/out] an object.  A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $sSearchString       - a string value. A String of text or a Regular Expression to Search for.
;                  $sReplaceString      - a string value. A String of text or a Regular Expression to replace any results with.
;                  $atFindFormat        - [in/out] an array of structs. Set to an empty array[0] to skip. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search". Array will not be modified.
;                  $atReplaceFormat     - [in/out] an array of structs. Set to an empty array[0] to skip. An Array of Formatting property values to replace any results with. Array will not be modified.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;				   @Error 1 @Extended 4 Return 0 = $sSearchString not a String.
;				   @Error 1 @Extended 5 Return 0 = $sReplaceString not a String.
;				   @Error 1 @Extended 6 Return 0 = $atFindFormat not an Array.
;				   @Error 1 @Extended 7 Return 0 = $atReplaceFormat not an Array.
;				   @Error 1 @Extended 8 Return 0 = First Element of $atFindFormat not an Object.
;				   @Error 1 @Extended 9 Return 0 = First Element of $atReplaceFormat not an Object.
;				   --Success--
;				   @Error 0 @Extended ? Return 1 = Success. Search and Replace was successful, @Extended set to number of replacements made.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_SearchDescriptorCreate, _LOWriter_DocFindAll, _LOWriter_DocFindNext,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocReplaceAllInRange, _LOWriter_FindFormatModifyAlignment,
;					_LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation,
;					_LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak,
;					_LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace,
;					_LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout,
;					_LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocReplaceAll(ByRef $oDoc, ByRef $oSrchDescript, $sSearchString, $sReplaceString, ByRef $atFindFormat, ByRef $atReplaceFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iReplacements

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSearchString) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sReplaceString) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If Not IsArray($atFindFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	If Not IsArray($atReplaceFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)

	If (UBound($atFindFormat) > 0) And Not IsObj($atFindFormat[0]) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
	$oSrchDescript.setSearchAttributes($atFindFormat)

	If (UBound($atReplaceFormat) > 0) And Not IsObj($atReplaceFormat[0]) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
	$oSrchDescript.setReplaceAttributes($atReplaceFormat)

	$oSrchDescript.SearchString = $sSearchString
	$oSrchDescript.ReplaceString = $sReplaceString

	$iReplacements = $oDoc.replaceAll($oSrchDescript)

	Return SetError($__LOW_STATUS_SUCCESS, $iReplacements, 1)

EndFunc   ;==>_LOWriter_DocReplaceAll

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocReplaceAllInRange
; Description ...: Replace all instances of a search within a selection. See Remarks.
; Syntax ........: _LOWriter_DocReplaceAllInRange(Byref $oDoc, Byref $oSrchDescript, Byref $oRange, $sSearchString, $sReplaceString, Byref $atFindFormat, Byref $atReplaceFormat)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $oRange              - [in/out] an object. A Range, such as a cursor with Data selected, to perform the search within.
;                  $sSearchString       - a string value. A String of text or a regular expression to search for.
;                  $sReplaceString      - a string value. A String of text or a regular expression to replace any results with.
;                  $atFindFormat        - [in/out] an array of structs. An Array of Formatting properties to search for, either by value or simply by existence, depending on the current setting of "Value Search".
;				   +						Set to an empty array[0] to skip. Array will not be modified.
;                  $atReplaceFormat     - [in/out] an array of structs. An Array of Formatting property values to replace any
;				   +						Set to an empty array[0] to skip. Array will not be modified. Not results with. Recommended for use with regular expressions, see remarks.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oSrchDescript not a Search Descriptor Object.
;				   @Error 1 @Extended 4 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 5 Return 0 = $oRange contains no selected Data.
;				   @Error 1 @Extended 6 Return 0 = $sSearchString not a String.
;				   @Error 1 @Extended 7 Return 0 = $sReplaceString not a String.
;				   @Error 1 @Extended 8 Return 0 = $atFindFormat not an Array.
;				   @Error 1 @Extended 9 Return 0 = $atReplaceFormat not an Array.
;				   @Error 1 @Extended 10 Return 0 = $atFindFormat is an Array but the First Element is not a Property Object.
;				   @Error 1 @Extended 11 Return 0 = $atReplaceFormat is an Array but the First Element is not a Property Object.
;				   @Error 1 @Extended 12 Return 0 = Search Styles is True, $atFindFormat and $atReplaceFormat arrays are empty, (Thus searching for Paragraph Styles by Name contained in the document) but $sReplaceString is set to a Paragraph Style that does not exist.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ViewCursor object.
;				   @Error 2 @Extended 2 Return 0 = Error creating backup of ViewCursor location and selection.
;				   @Error 2 @Extended 3 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;				   @Error 2 @Extended 4 Return 0 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error converting Regular Expression String.
;				   @Error 3 @Extended 2 Return 0 = Error performing FindAllInRange Function.
;				   --Success--
;				   @Error 0 @Extended ? Return 1 = Success. Search and Replace was successful, number of replacements returned in @Extended.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Libre Office does not offer a Function to call to replace only results within a selection,
;						consequently I have had to create my own. This function uses the "FindAllInRange" function,
;						so any errors with Find/Replace formatting causing deletions will cause problems here. As best
;						as I can tell all options for find and replace should be available, Formatting, Paragraph styles
;						etc. How I created this function to still accept Regular Expressions is I use Libre's FindAll
;						command, modified by my FindAllInRange function. I then ran into another problem, as my next step
;						was to use AutoIt's RegExpReplace function to perform the replacement, but some replacements don't
;						work as expected. To Fix this I have created two versions of Regular Expression replacement, the
;						first way is only implemented if $atReplaceFormat is skipped using an empty array. I use an
;						ExecutionHelper to execute the Find and replace command, however this method doesn't accept formatting
;						for find and replace. So I developed my second method, which accepts formatting, and uses AutoIt's
;						RegExpReplace function to "Search" the resulting matched Strings and replace it, then I set the new
;						string to that result. However I have had to create a separate function to convert the ReplaceString
;						to be compatible with AutoIt's Regular Expression formatting. A BackSlash (\) must be doubled(\\) in
;						order to be literally inserted, at the beginning of the conversion process all double Backslashes
;						are replaced with a specific flag to aid in identifying commented and non-commented keywords
;						(\n, \t, & etc.), after the  conversion process the special flag is replaced again with the double
;						Backslashes, this should not cause any issues, \n (new Paragraph) in L.O. RegExp. formatting is
;						replaced with @CR, unless the Backslash is doubled (\\n), then \n becomes literal, \t (Tab) in L.O.
;						format is replaced with @Tab, and &(Find Result/BackReference) is replaced with $0 which means insert
;						the entire found string at that position, To insert a regular "&" character, comment it with a
;						Backslash, \&. As with LibreOffice, this function should still accept BackReferences ($0-9 or \0-9).
;						However I have found certain problems with some of the expressions still not working, such as
;						$ (end of paragraph mark) not replacing correctly because Autoit uses @CRLF for its newline marks, and
;						Libre uses @CR for a paragraph and @LF for a soft newline.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_SearchDescriptorCreate, _LOWriter_DocFindAll, _LOWriter_DocFindNext,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocReplaceAll, _LOWriter_FindFormatModifyAlignment,
;					_LOWriter_FindFormatModifyEffects, _LOWriter_FindFormatModifyFont, _LOWriter_FindFormatModifyHyphenation,
;					_LOWriter_FindFormatModifyIndent, _LOWriter_FindFormatModifyOverline, _LOWriter_FindFormatModifyPageBreak,
;					_LOWriter_FindFormatModifyPosition, _LOWriter_FindFormatModifyRotateScaleSpace,
;					_LOWriter_FindFormatModifySpacing, _LOWriter_FindFormatModifyStrikeout,
;					_LOWriter_FindFormatModifyTxtFlowOpt, _LOWriter_FindFormatModifyUnderline.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocReplaceAllInRange(ByRef $oDoc, ByRef $oSrchDescript, ByRef $oRange, $sSearchString, $sReplaceString, ByRef $atFindFormat, ByRef $atReplaceFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aoResults[0]
	Local $atArgs[7]
	Local Const $LOW_SEARCHFLAG_ABSOLUTE = 1, $LOW_SEARCHFLAG_REGEXP = 2, $LOW_SEARCHFLAG_REPLACE_ALL = 3, $LOW_SEARCHFLAG_SELECTION = 2048
	Local $oViewCursor, $oViewCursorBackup, $oServiceManager, $oDispatcher
	Local $iResults
	Local $bFormat = False

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsObj($oRange) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If ($oRange.IsCollapsed()) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If Not IsString($sSearchString) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	If Not IsString($sReplaceString) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
	If Not IsArray($atFindFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
	If Not IsArray($atReplaceFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)

	If (UBound($atFindFormat) > 0) And Not IsObj($atFindFormat[0]) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
	If (UBound($atReplaceFormat) > 0) And Not IsObj($atReplaceFormat[0]) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 11, 0)
	If (UBound($atReplaceFormat) > 0) Then $bFormat = True

	; If Find/Replace using a Regular expression is True, and replace formatting is set, convert the regular expressions for my
	; alternate replacement function to use.
	If ($oSrchDescript.SearchRegularExpression() = True) And ($bFormat = True) Then __LOWriter_RegExpConvert($sReplaceString)
	If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	$aoResults = _LOWriter_DocFindAllInRange($oDoc, $oSrchDescript, $sSearchString, $atFindFormat, $oRange)
	$iResults = @extended
	If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0) ; Error performing search

	If ($oSrchDescript.SearchRegularExpression() = True) Then

		If ($bFormat = True) Then

			For $i = 0 To $iResults - 1
				$aoResults[$i].setString(StringRegExpReplace($aoResults[$i].getString(), $sSearchString, $sReplaceString))
				If ($bFormat = True) Then
					For $j = 0 To UBound($atReplaceFormat) - 1
						$aoResults[$i].setPropertyValue($atReplaceFormat[$j].Name(), $atReplaceFormat[$j].Value())
					Next
				EndIf

				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
			Next

		Else ;No Replacement formatting, use UNO Execute method instead.

			$oViewCursor = $oDoc.CurrentController.getViewCursor()
			If Not IsObj($oViewCursor) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

			; Backup the ViewCursor location and selection.
			$oViewCursorBackup = $oDoc.Text.createTextCursorByRange($oViewCursor)
			If Not IsObj($oViewCursorBackup) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

			; Move the View Cursor to the input range.
			$oViewCursor.gotoRange($oRange, False)

			$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
			If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

			$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
			If Not IsObj($oDispatcher) Then Return SetError($__LOW_STATUS_INIT_ERROR, 4, 0)

			$atArgs[0] = __LOWriter_SetPropertyValue("SearchItem.Backward", $oSrchDescript.SearchBackwards())
			$atArgs[1] = __LOWriter_SetPropertyValue("SearchItem.AlgorithmType", $LOW_SEARCHFLAG_ABSOLUTE)
			$atArgs[2] = __LOWriter_SetPropertyValue("SearchItem.SearchFlags", $LOW_SEARCHFLAG_SELECTION)
			$atArgs[3] = __LOWriter_SetPropertyValue("SearchItem.SearchString", $sSearchString)
			$atArgs[4] = __LOWriter_SetPropertyValue("SearchItem.ReplaceString", $sReplaceString)
			$atArgs[5] = __LOWriter_SetPropertyValue("SearchItem.Command", $LOW_SEARCHFLAG_REPLACE_ALL)
			$atArgs[6] = __LOWriter_SetPropertyValue("SearchItem.AlgorithmType2", $LOW_SEARCHFLAG_REGEXP)

			$oDispatcher.executeDispatch($oDoc.CurrentController, ".uno:ExecuteSearch", "", 0, $atArgs)

			; Restore the ViewCursor to its previous location.
			$oViewCursor.gotoRange($oViewCursorBackup, False)

		EndIf

	ElseIf ($oSrchDescript.SearchStyles() = True) And ((UBound($atFindFormat) = 0) And (UBound($atReplaceFormat) = 0)) Then ; If Style Search is active and no formatting is set, then search and replace Paragraph style.
		If Not _LOWriter_ParStyleExists($oDoc, $sReplaceString) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 12, 0)
		For $i = 0 To $iResults - 1
			$aoResults[$i].ParaStyleName = $sReplaceString

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	Else
		For $i = 0 To $iResults - 1
			$aoResults[$i].setString(StringReplace($aoResults[$i].getString(), $sSearchString, $sReplaceString))
			If ($bFormat = True) Then
				For $k = 0 To UBound($atReplaceFormat) - 1
					$aoResults[$i].setPropertyValue($atReplaceFormat[$k].Name(), $atReplaceFormat[$k].Value())
				Next
			EndIf

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, $iResults, 1)
EndFunc   ;==>_LOWriter_DocReplaceAllInRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocSave
; Description ...: Save any changes made to a Document.
; Syntax ........: _LOWriter_DocSave(Byref $oDoc)
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Document is ReadOnly or Document has no save location, try SaveAs.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Document Successfully saved.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDoc.hasLocation Or $oDoc.isReadOnly Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	$oDoc.store()
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocSaveAs
; Description ...: Save a Document with the specified file name to the path specified with any parameters called.
; Syntax ........: _LOWriter_DocSaveAs(Byref $oDoc, $sFilePath[, $sFilterName = ""[, $bOverwrite = Null[, $sPassword = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFilePath           - a string value. Full path to save the document to, including Filename and extension.
;                  $sFilterName         - [optional] a string value. Default is "". Filter name. "" (blank string), Filter is chosen automatically based on the file extension.
;				   +						If no extension is present, or if not matched to the list of extensions in this UDF, the .odt extension is used instead, with the filter name of "writer8".
;                  $bOverwrite          - [optional] a boolean value. Default is Null. If True, the existing file will be overwritten.
;                  $sPassword           - [optional] a string value. Default is Null. Sets a password for the document. (Not all file formats can have a Password set). "" (blank string) = No Password. Null also sets no password.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFilePath Not a String.
;				   @Error 1 @Extended 3 Return 0 = $sFilterName Not a String.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite Not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sPassword Not a String.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error Converting Path to/from L.O. URL
;				   @Error 3 @Extended 2 Return 0 = Error retrieving FilterName.
;				   @Error 3 @Extended 3 Return 0 = Error setting FilterName Property
;				   @Error 3 @Extended 4 Return 0 = Error setting Overwrite Property
;				   @Error 3 @Extended 5 Return 0 = Error setting Password Property
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Successfully Saved the document. Returning document save path.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFilterName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$sFilePath = _LOWriter_PathConvert($sFilePath, $LOW_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	If ($sFilterName = "") Or ($sFilterName = " ") Then $sFilterName = __LOWriter_FilterNameGet($sFilePath)
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
	$aProperties[0] = __LOWriter_SetPropertyValue("FilterName", $sFilterName)
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 3, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOWriter_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	If $sPassword <> Null Then
		If Not IsString($sPassword) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOWriter_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	$oDoc.storeAsURL($sFilePath, $aProperties)

	$sSavePath = _LOWriter_PathConvert($sFilePath, $LOW_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOWriter_DocSaveAs

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocToFront
; Description ...: Bring the called document to the front of the windows.
; Syntax ........: _LOWriter_DocToFront(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Windows was successfully brought to the front of the open windows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If minimized, the document is restored and brought to the front of the viewable pages. Generally only brings
;					the document to the front of other Libre Office windows.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocToFront(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oDoc.CurrentController.Frame.ContainerWindow.toFront()
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocToFront

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndo
; Description ...: Perform one Undo action for a document.
; Syntax ........: _LOWriter_DocUndo(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Document does not have an undo action to perform.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Successfully performed an undo action.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocUndoIsPossible, _LOWriter_DocUndoGetAllActionTitles, _LOWriter_DocUndoCurActionTitle,
;					_LOWriter_DocRedo
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocUndo(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isUndoPossible()) Then
		$oDoc.UndoManager.Undo()
		Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
	Else
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOWriter_DocUndo

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoCurActionTitle
; Description ...: Retrieve the current Undo action Title.
; Syntax ........: _LOWriter_DocUndoCurActionTitle(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Document does not have an undo action available.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Returns the current available undo action Title in String format.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isUndoPossible()) Then
		Return SetError($__LOW_STATUS_SUCCESS, 1, $oDoc.UndoManager.getCurrentUndoActionTitle())
	Else
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

EndFunc   ;==>_LOWriter_DocUndoCurActionTitle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoGetAllActionTitles
; Description ...: Retrieve all available Undo action Titles.
; Syntax ........: _LOWriter_DocUndoGetAllActionTitles(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Document does not have any undo actions available.
;				   --Success--
;				   @Error 0 @Extended 0 Return Array = Returns all available undo action Titles in an array of Strings.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isUndoPossible()) Then
		Return SetError($__LOW_STATUS_SUCCESS, 1, $oDoc.UndoManager.getAllUndoActionTitles())
	Else
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOWriter_DocUndoGetAllActionTitles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocUndoIsPossible
; Description ...: Test whether a Undo action is available to perform for a document.
; Syntax ........: _LOWriter_DocUndoIsPossible(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If the document has an undo action to perform, True is returned, else False.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 1, $oDoc.UndoManager.isUndoPossible())
EndFunc   ;==>_LOWriter_DocUndoIsPossible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocViewCursorGetPosition
; Description ...: Retrieve View Cursor position in Micrometers.
; Syntax ........: _LOWriter_DocViewCursorGetPosition(Byref $oCursor)
; Parameters ....: $oCursor             - [in/out] an object. A View Cursor Object returned by _LOWriter_DocGetViewCursor function.
; Return values .: Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not a View Cursor.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error determining Cursor type.
;				   --Success--
;				   @Error 0 @Extended ? Return Integer = Success. Current Cursor Coordinate position relative to the top-left position of the first page of the document returned.
;				   +	@Extended is the "X" coordinate, and Return value is the "Y" Coordinate. In Micrometers.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_CursorMove, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DocViewCursorGetPosition(ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType

	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType <> $LOW_CURTYPE_VIEW_CURSOR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	Return SetError($__LOW_STATUS_SUCCESS, $oCursor.getPosition().X(), $oCursor.getPosition().Y())

EndFunc   ;==>_LOWriter_DocViewCursorGetPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocVisible
; Description ...: Set or retrieve the current visibility of a document.
; Syntax ........: _LOWriter_DocVisible(Byref $oDoc[, $bVisible = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the document is visible, else if false, document becomes invisible.
; Return values .: Success: 1 or Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Error setting $bVisible.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. $bVisible successfully set.
;				   @Error 0 @Extended 1 Return Boolean = Success. Returning current visibility state of the Document, True if visible, false if invisible.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If ($bVisible = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oDoc.CurrentController.Frame.ContainerWindow.isVisible())

	If Not IsBool($bVisible) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Frame.ContainerWindow.Visible = $bVisible
	$iError = ($oDoc.CurrentController.Frame.ContainerWindow.isVisible() = $bVisible) ? 0 : 1

	Return ($iError = 0) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0)
EndFunc   ;==>_LOWriter_DocVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DocZoom
; Description ...: Modify the zoom value for a document.
; Syntax ........: _LOWriter_DocZoom(Byref $oDoc[, $iZoom = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iZoom               - [optional] an integer value. Default is Null. The zoom percentage. Min. 20%, Max 600%.
; Return values .: Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iZoom not an Integer, less than 20 or greater than 600.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.DispatchHelper" Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 	0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |							1 = Error setting $iZoom
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = $iZoom set successfully.
;				   @Error 0 @Extended 1 Return Integer = $iZoom set to null, returning current zoom value.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iZoom = Null) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $oDoc.CurrentController.ViewSettings.ZoomValue())

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$oDispatcher = $oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
	If Not IsObj($oDispatcher) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	If Not __LOWriter_IntIsBetween($iZoom, 20, 600) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$aArgs[0] = __LOWriter_SetPropertyValue("Zoom.Value", $iZoom)
	$aArgs[1] = __LOWriter_SetPropertyValue("Zoom.ValueSet", 28703)
	$aArgs[2] = __LOWriter_SetPropertyValue("Zoom.Type", 0)

	$oDispatcher.executeDispatch($oDoc.CurrentController, ".uno:Zoom", "", 0, $aArgs)
	$iError = ($oDoc.CurrentController.ViewSettings.ZoomValue() = $iZoom) ? $iError : BitOR($iError, 1)

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DocZoom
