#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

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
; _LOCalc_DocConnect
; _LOCalc_DocCreate
; _LOCalc_DocExport
; _LOCalc_DocGetName
; _LOCalc_DocGetPath
; _LOCalc_DocHasPath
; _LOCalc_DocHasSheetName
; _LOCalc_DocIsActive
; _LOCalc_DocIsModified
; _LOCalc_DocIsReadOnly
; _LOCalc_DocMaximize
; _LOCalc_DocMinimize
; _LOCalc_DocOpen
; _LOCalc_DocPosAndSize
; _LOCalc_DocRedo
; _LOCalc_DocRedoCurActionTitle
; _LOCalc_DocRedoGetAllActionTitles
; _LOCalc_DocRedoIsPossible
; _LOCalc_DocSave
; _LOCalc_DocSaveAs
; _LOCalc_DocToFront
; _LOCalc_DocUndo
; _LOCalc_DocUndoCurActionTitle
; _LOCalc_DocUndoGetAllActionTitles
; _LOCalc_DocUndoIsPossible
; _LOCalc_DocVisible
; _LOCalc_DocZoom
; ===============================================================================================================================


; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocClose
; Description ...: Close an existing Calc Document, returning its save path if applicable.
; Syntax ........: _LOCalc_DocClose(ByRef $oDoc[, $bSaveChanges = True[, $sSaveName = ""[, $bDeliverOwnership = True]]])
; Parameters ....: $oDoc           		- [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bSaveChanges        - [optional] a boolean value. Default is True. If true, saves changes if any were made before closing. See remarks.
;                  $sSaveName           - [optional] a string value. Default is "". The file name to save the file as, if the file hasn't been saved before. See Remarks.
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
;				   @Error 3 @Extended 3 Return 0 = Error while setting Filter Name properties.
;				   --Success--
;				   @Error 0 @Extended 1 Return String = Success, Document was successfully closed, and was saved to the returned file Path.
;				   @Error 0 @Extended 2 Return String = Success, Document was successfully closed, document's changes were saved to its existing location.
;				   @Error 0 @Extended 3 Return String = Success, Document was successfully closed, document either had no changes to save, or $bSaveChanges was set to False.
;				   +			If document had a save location, or if document was saved to a location, it is returned, else an empty string is returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bSaveChanges is true and the document hasn't been saved yet, the document is saved to the desktop.
;					If $sSaveName is undefined, it is saved as an .odt document to the desktop, named
;					Year-Month-Day_Hour-Minute-Second.ods. $sSaveName may be a name only without an extension, in which case the file will be saved
;					in .ods format. Or you may define your own format by including an extension, such as "Test.xlsx"
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

		$sSavePath = _LOCalc_PathConvert($sSavePath & $sSaveName, 1)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If $sFilterName = "" Then $sFilterName = __LOCalc_FilterNameGet($sSavePath)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$aArgs[0] = __LOCalc_SetPropertyValue("FilterName", $sFilterName)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	EndIf

	If ($bSaveChanges = True) Then

		If $oDoc.hasLocation() Then
			$oDoc.store()
			$sDocPath = _LOCalc_PathConvert($oDoc.getURL(), $LOC_PATHCONV_PCPATH_RETURN)
			$oDoc.Close($bDeliverOwnership)
			Return SetError($__LO_STATUS_SUCCESS, 2, $sDocPath)
		Else
			$oDoc.storeAsURL($sSavePath, $aArgs)
			$oDoc.Close($bDeliverOwnership)
			Return SetError($__LO_STATUS_SUCCESS, 1, _LOCalc_PathConvert($sSavePath, $LOC_PATHCONV_PCPATH_RETURN))
		EndIf

	EndIf

	If $oDoc.hasLocation() Then $sDocPath = _LOCalc_PathConvert($oDoc.getURL(), $LOC_PATHCONV_PCPATH_RETURN)
	$oDoc.Close($bDeliverOwnership)
	Return SetError($__LO_STATUS_SUCCESS, 3, $sDocPath)

EndFunc   ;==>_LOCalc_DocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocConnect
; Description ...: Connect to an already opened instance of LibreOffice Calc.
; Syntax ........: _LOCalc_DocConnect($sFile[, $bConnectCurrent = False[, $bConnectAll = False]])
; Parameters ....: $sFile               - a string value. A Full or partial file path, or a full or partial file name. See remarks. Can be an empty string if $bConnectAll or $bConnectCurrent is True.
;                  $bConnectCurrent     - [optional] a boolean value. Default is False. If True, returns the currently active, or last active Document, unless it is not a Calc Document.
;                  $bConnectAll         - [optional] a boolean value. Default is False. If True, returns an array containing all open LibreOffice Calc Documents. See remarks.
; -Return values .: Success: Object or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sFile not a string.
;				   @Error 1 @Extended 2 Return 0 = $bConnectCurrent not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bConnectAll not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating ServiceManager object.
;				   @Error 2 @Extended 2 Return 0 = Error creating Desktop object.
;				   @Error 2 @Extended 3 Return 0 = Error creating enumeration of open documents.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error converting path to Libre Office URL.
;				   --Document Errors--
;				   @Error 5 @Extended 1 Return 0 = No matches found.
;				   @Error 5 @Extended 2 Return 0 = Current Component not a Calc Document.
;				   @Error 5 @Extended 3 Return 0 = No open Libre Office documents found.
;				   --Success--
;				   @Error 0 @Extended 1 Return Object =  Success, The Object for the current, or last active document is returned.
;				   @Error 0 @Extended 2 Returns Array =  Success, An Array of all open LibreOffice Calc documents is returned. See remarks.
;				   @Error 0 @Extended 3 Return Object =  Success, The Object for the document with matching URL is returned.
;				   @Error 0 @Extended 4 Return Object =  Success, The Object for the document with matching Title is returned.
;				   @Error 0 @Extended 5 Return Object =  Success, A partial Title or Path search found only one match, returning the Object for the found document.
;				   @Error 0 @Extended 6 Return Array =  Success, An Array of all matching Libre Text documents from a partial Title or Path search. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:  $sFile can be either the full Path (Name and extension included; i.e: C:\file\Test.ods Or file:///C:/file/Test.ods) of the document, or the full Title with extension, (i.e: Test.ods), or a partial file path (i.e: file1\file2\Test Or file1\file2 Or file1/file2/ etc.), or a partial name (i.e: test, would match Test1.ods, Test2.xlsx etc.).
;					Partial file path searches and file name searches, as well as the connect all option, return arrays with three columns per result. ($aArray[0][3]. each result is stored in a separate row;
;				    Row 1, Column 0 contain the Object variable for that document. i.e. $aArray[0][0] = $oDoc
;					Row 1, Column 1 contains the Document's full title and extension. i.e. $aArray[0][1] = This Test File.xlsx
;					Row 1, Column 2 contains the document's full file path. i.e. $aArray[0][2] = C:\Folder1\Folder2\This Test File.xlsx
;					Row 2, Column 0 contain the Object variable for the next document. And so on. i.e. $aArray[1][0] = $oDoc2
; Related .......: _LOCalc_DocOpen, _LOCalc_DocClose, _LOCalc_DocCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocConnect($sFile, $bConnectCurrent = False, $bConnectAll = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local Const $STR_STRIPLEADING = 1
	Local $aoConnectAll[1], $aoPartNameSearch[1]
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop
	Local $sServiceName = "com.sun.star.sheet.SpreadsheetDocument"

	If Not IsString($sFile) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bConnectCurrent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bConnectAll) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not $oDesktop.getComponents.hasElements() Then Return SetError($__LO_STATUS_DOC_ERROR, 3, 0) ; no L.O open
	$oEnumDoc = $oDesktop.getComponents.createEnumeration()
	If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If $bConnectCurrent Then
		$oDoc = $oDesktop.currentComponent()
		Return ($oDoc.supportsService($sServiceName)) ? (SetError($__LO_STATUS_SUCCESS, 1, $oDoc)) : (SetError($__LO_STATUS_DOC_ERROR, 2, 0))
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
				$aoConnectAll[$iCount][2] = _LOCalc_PathConvert($oDoc.getURL(), $LOC_PATHCONV_PCPATH_RETURN)
				$iCount += 1
			EndIf
			Sleep(10)
		WEnd
		Return SetError($__LO_STATUS_SUCCESS, 2, $aoConnectAll)

	EndIf

	$sFile = StringStripWS($sFile, $STR_STRIPLEADING)
	If StringInStr($sFile, "\") Then $sFile = _LOCalc_PathConvert($sFile, $LOC_PATHCONV_OFFICE_RETURN) ; Convert to L.O File path.
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If StringInStr($sFile, "file:///") Then ; URL/Path and Name search

		While $oEnumDoc.hasMoreElements()
			$oDoc = $oEnumDoc.nextElement()

			If ($oDoc.getURL() == $sFile) Then Return SetError($__LO_STATUS_SUCCESS, 3, $oDoc) ; Match
		WEnd
		Return SetError($__LO_STATUS_DOC_ERROR, 1, 0) ; no match

	Else
		If Not StringInStr($sFile, "/") And StringInStr($sFile, ".") Then ; Name with extension only search
			While $oEnumDoc.hasMoreElements()
				$oDoc = $oEnumDoc.nextElement()
				If StringInStr($oDoc.Title, $sFile) Then Return SetError($__LO_STATUS_SUCCESS, 4, $oDoc) ; Match
			WEnd
			Return SetError($__LO_STATUS_DOC_ERROR, 1, 0) ; no match
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
					$aoPartNameSearch[$iCount][2] = _LOCalc_PathConvert($oDoc.getURL, $LOC_PATHCONV_PCPATH_RETURN)
					$iCount += 1
				EndIf
			Else
				If StringInStr($oDoc.Title, $sFile) Then
					ReDim $aoPartNameSearch[$iCount + 1][3]
					$aoPartNameSearch[$iCount][0] = $oDoc
					$aoPartNameSearch[$iCount][1] = $oDoc.Title
					$aoPartNameSearch[$iCount][2] = _LOCalc_PathConvert($oDoc.getURL, $LOC_PATHCONV_PCPATH_RETURN)
					$iCount += 1
				EndIf
			EndIf

		WEnd
		If IsString($aoPartNameSearch[0][1]) Then
			If (UBound($aoPartNameSearch) = 1) Then
				Return SetError($__LO_STATUS_SUCCESS, 5, $aoPartNameSearch[0][0]) ; matches
			Else
				Return SetError($__LO_STATUS_SUCCESS, 6, $aoPartNameSearch) ; matches
			EndIf

		Else
			Return SetError($__LO_STATUS_DOC_ERROR, 1, 0) ; no match
		EndIf

	EndIf

EndFunc   ;==>_LOCalc_DocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocCreate
; Description ...: Open a new Libre Office Calc Document or Connect to an existing blank, unsaved, writable document.
; Syntax ........: _LOCalc_DocCreate([$bForceNew = True[, $bHidden = False]])
; Parameters ....: $bForceNew		- [optional] a boolean value. Default is True. If True, force opening a new Calc Document instead of checking for a usable blank.
;				   $bHidden			- [optional] a boolean value. Default is False. If True opens the new document invisible or changes the existing document to invisible.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $bForceNew not a Boolean.
;				   @Error 1 @Extended 2 Return 0 = $bHidden not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failure Creating Object com.sun.star.ServiceManager.
;				   @Error 2 @Extended 2 Return 0 = Failure Creating Object com.sun.star.frame.Desktop.
;				   @Error 2 @Extended 3 Return 0 = Failed to enumerate available documents.
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
; Related .......: _LOCalc_DocOpen, _LOCalc_DocClose, _LOCalc_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocCreate($bForceNew = True, $bHidden = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iURLFrameCreate = 8 ;frame will be created if not found
	Local $aArgs[1]
	Local $iError = 0
	Local $oServiceManager, $oDesktop, $oDoc, $oEnumDoc
	Local $sServiceName = "com.sun.star.sheet.SpreadsheetDocument"

	If Not IsBool($bForceNew) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$aArgs[0] = __LOCalc_SetPropertyValue("Hidden", $bHidden)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
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
					And Not ($oDoc.hasLocation() And $oDoc.isReadOnly()) And ($oDoc.WordCount() = 0) Then
				$oDoc.CurrentController.Frame.ContainerWindow.Visible = ($bHidden) ? (False) : (True) ; opposite value of $bHidden.
				$iError = ($oDoc.CurrentController.Frame.isHidden() = $bHidden) ? ($iError) : (BitOR($iError, 1))
				Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))
			EndIf
		WEnd
	EndIf

	If Not IsObj($aArgs[0]) Then Return $iError = BitOR($iError, 1)
	$oDoc = $oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", $iURLFrameCreate, $aArgs)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOCalc_DocCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocExport
; Description ...: Export a Document with the specified file name to the path specified, with any parameters used.
; Syntax ........: _LOCalc_DocExport(ByRef $oDoc, $sFilePath[, $bSamePath = False[, $sFilterName = ""[, $bOverwrite = Null[, $sPassword = Null]]]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sFilePath      - a string value. Full path to save the document to, including Filename and extension. See Remarks.
;                  $bSamePath      - [optional] a boolean value. Default is False. If True, uses the path of the current document to export to. See Remarks
;                  $sFilterName    - [optional] a string value. Default is "". Filter name. If set to "" (blank string), Filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .ods extension is used instead, with the filter name of "calc8".
;                  $bOverwrite     - [optional] a boolean value. Default is Null. If True, file will be overwritten.
;                  $sPassword      - [optional] a string value. Default is Null. Password String to set for the document. (Not all file formats can have a Password set). "" (blank string) or Null = No Password.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFilePath not a String.
;				   @Error 1 @Extended 3 Return 0 = $bSamePath not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $sFilterName not a String.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sPassword not a String.
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
;				   If $bSamePath is set to True, the same save path as the current document is used. You must still fill in "sFilePath" with the desired File Name and new extension, but you do not need to enter the file path.
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
			If StringInStr($sFilePath, "\") Then $sFilePath = _LOCalc_PathConvert($sFilePath, $LOC_PATHCONV_OFFICE_RETURN) ; Convert to L.O. URL
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
			$sFilePath = $sOriginalPath & $sFilePath ; combine the path with the new name.
		Else
			Return SetError($__LO_STATUS_DOC_ERROR, 1, 0)
		EndIf
	EndIf

	If Not $bSamePath Then $sFilePath = _LOCalc_PathConvert($sFilePath, $LOC_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sFilterName = "") Or ($sFilterName = " ") Then $sFilterName = __LOCalc_FilterNameGet($sFilePath, True)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$aProperties[0] = __LOCalc_SetPropertyValue("FilterName", $sFilterName)
	If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOCalc_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0)
	EndIf

	If ($sPassword <> Null) Then
		If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOCalc_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0)
	EndIf

	$oDoc.storeToURL($sFilePath, $aProperties)

	$sSavePath = _LOCalc_PathConvert($sFilePath, $LOC_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOCalc_DocExport

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocGetName
; Description ...: Retrieve the document's name.
; Syntax ........: _LOCalc_DocGetName(ByRef $oDoc[, $bReturnFull = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bReturnFull         - [optional] a boolean value. Default is False. If True, the full window title is returned, such as is used by AutoIt window related functions.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bReturnFull not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returns the document's current Name/Title
;				   @Error 0 @Extended 1 Return String = Success. Returns the document's current Window Title, which includes the document name and usually: "-LibreOffice Calc".
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bReturnLibreURL     - [optional] a boolean value. Default is False. If True, returns a path in Libre Office URL format, else false returns a regular Windows path.
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
; Related .......: _LOCalc_PathConvert
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

	If ($bReturnLibreURL = True) Then
		$sPath = $oDoc.URL()
	Else
		$sPath = $oDoc.URL()
		$sPath = _LOCalc_PathConvert($sPath, $LOC_PATHCONV_PCPATH_RETURN)
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return ($bReturnLibreURL = True) ? (SetError($__LO_STATUS_SUCCESS, 1, $sPath)) : (SetError($__LO_STATUS_SUCCESS, 0, $sPath))
EndFunc   ;==>_LOCalc_DocGetPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocHasPath
; Description ...: Returns whether a document has been saved to a location already or not.
; Syntax ........: _LOCalc_DocHasPath(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
Func _LOCalc_DocHasPath(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc.hasLocation())
EndFunc   ;==>_LOCalc_DocHasPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocHasSheetName
; Description ...: Check whether a Calc document has a Sheet with a specific name.
; Syntax ........: _LOCalc_DocHasSheetName(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sName               - a string value. The sheet name to check for.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Sheets Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. If the document contains a Sheet matching $sName, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocHasSheetName(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSheets

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSheets = $oDoc.Sheets()
	If Not IsObj($oSheets) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $oSheets.hasByName($sName) Then Return SetError($__LO_STATUS_Success, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOCalc_DocHasSheetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocIsActive
; Description ...: Tests if called document is the active document of other Libre windows.
; Syntax ........: _LOCalc_DocIsActive(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. Returns True if document is the currently active Libre window. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note, this does NOT test if the document is the current active window in Windows, it only tests if the document is the current active document among other Libre Office documents.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
; Syntax ........: _LOCalc_DocOpen($sFilePath[, $bConnectIfOpen = True[, $bHidden = False[, $bReadOnly = False[, $sPassword = ""[, $bLoadAsTemplate = False[, $sFilterName = ""]]]]]])
; Parameters ....: $sFilePath           - a string value. Full path and filename of the file to be opened.
;                  $bConnectIfOpen      - [optional] a boolean value. Default is True(Connect). Whether to connect to the requested document if it is already open. See remarks.
;                  $bHidden             - [optional] a boolean value. Default is Null. If true, opens the document invisibly.
;                  $bReadOnly           - [optional] a boolean value. Default is Null. If true, opens the document as read-only.
;                  $sPassword           - [optional] a string value. Default is Null. The password that was used to read-protect the document, if any.
;                  $bLoadAsTemplate     - [optional] a boolean value. Default is Null. If true, opens the document as a Template, i.e. an untitled copy of the specified document is made instead of modifying the original document.
;                  $sFilterName         - [optional] a string value. Default is Null. Name of a LibreOffice filter to use to load the specified document. LibreOffice automatically selects which to use by default.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sFilePath not string, or file not found.
;				   @Error 1 @Extended 2 Return 0 = Error converting file path to URL path.
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
; Remarks .......: Note: any parameters (Hidden, template etc.,) will not be applied when connecting to a document.
; Related .......: _LOCalc_DocCreate, _LOCalc_DocClose, _LOCalc_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocOpen($sFilePath, $bConnectIfOpen = True, $bHidden = Null, $bReadOnly = Null, $sPassword = Null, $bLoadAsTemplate = Null, $sFilterName = Null)
	Local Const $iURLFrameCreate = 8 ;frame will be created if not found
	Local $iError = 0
	Local $oDoc, $oServiceManager, $oDesktop
	Local $aoProperties[0]
	Local $vProperty
	Local $sFileURL

	Local $oComError = ObjEvent("AutoIt.Error", "__LOCalc_InternalComErrorHandler")

	If Not IsString($sFilePath) Or Not FileExists($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	$sFileURL = _LOCalc_PathConvert($sFilePath, $LOC_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bConnectIfOpen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not __LOCalc_VarsAreNull($bHidden, $bReadOnly, $sPassword, $bLoadAsTemplate, $sFilterName) Then

		If ($bHidden <> Null) Then
			If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
			$vProperty = __LOCalc_SetPropertyValue("Hidden", $bHidden)
			If @error Then $iError = BitOR($iError, 1)
			If Not BitAND($iError, 1) Then __LOCalc_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($bReadOnly <> Null) Then
			If Not IsBool($bReadOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
			$vProperty = __LOCalc_SetPropertyValue("ReadOnly", $bReadOnly)
			If @error Then $iError = BitOR($iError, 2)
			If Not BitAND($iError, 2) Then __LOCalc_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($sPassword <> Null) Then
			If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
			$vProperty = __LOCalc_SetPropertyValue("Password", $sPassword)
			If @error Then $iError = BitOR($iError, 4)
			If Not BitAND($iError, 4) Then __LOCalc_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($bLoadAsTemplate <> Null) Then
			If Not IsBool($bLoadAsTemplate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
			$vProperty = __LOCalc_SetPropertyValue("AsTemplate", $bLoadAsTemplate)
			If @error Then $iError = BitOR($iError, 8)
			If Not BitAND($iError, 8) Then __LOCalc_AddTo1DArray($aoProperties, $vProperty)
		EndIf

		If ($sFilterName <> Null) Then
			If Not IsString($sFilterName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
			$vProperty = __LOCalc_SetPropertyValue("FilterName", $sFilterName)
			If @error Then $iError = BitOR($iError, 16)
			If Not BitAND($iError, 16) Then __LOCalc_AddTo1DArray($aoProperties, $vProperty)
		EndIf

	EndIf

	$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
	If StringInStr($oComError.description, """type detection failed""") And $bConnectIfOpen Then
		ReDim $aoProperties[0]
		$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))
	EndIf

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)
	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOCalc_DocOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocPosAndSize
; Description ...: Reposition and resize a document window.
; Syntax ........: _LOCalc_DocPosAndSize(ByRef $oDoc[, $iX = Null[, $iY = Null[, $iWidth = Null[, $iHeight = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
;				   |								1 = Error setting $iX
;				   |								2 = Error setting $iY
;				   |								4 = Error setting $iWidth
;				   |								8 = Error setting $iHeight
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note, X & Y, on my computer at least, seem to go no lower than 8(X) and 30(Y), if you enter lower than this, it will cause a "property setting Error".
;				   If you want more accurate functionality, use the "WinMove" AutoIt function.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
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
	If Not IsObj($tWindowSize) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If __LOCalc_VarsAreNull($iX, $iY, $iWidth, $iHeight) Then
		__LOCalc_ArrayFill($aiWinPosSize, $tWindowSize.X(), $tWindowSize.Y(), $tWindowSize.Width(), $tWindowSize.Height())
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
	If Not IsObj($tWindowSize) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$iError = ($iX = Null) ? ($iError) : (($tWindowSize.X() = $iX) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iY = Null) ? ($iError) : (($tWindowSize.Y() = $iY) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iWidth = Null) ? ($iError) : (($tWindowSize.Width() = $iWidth) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iHeight = Null) ? ($iError) : (($tWindowSize.Height() = $iHeight) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc   ;==>_LOCalc_DocPosAndSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocRedo
; Description ...: Perform one Redo action for a document.
; Syntax ........: _LOCalc_DocRedo(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
; Name ..........: _LOCalc_DocRedoCurActionTitle
; Description ...: Retrieve the current Redo action Title.
; Syntax ........: _LOCalc_DocRedoCurActionTitle(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Document does not have a redo action available.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Returns the current available redo action Title as a String.
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

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isRedoPossible()) Then
		Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.UndoManager.getCurrentRedoActionTitle())
	Else
		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOCalc_DocRedoCurActionTitle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocRedoGetAllActionTitles
; Description ...: Retrieve all available Redo action Titles.
; Syntax ........: _LOCalc_DocRedoGetAllActionTitles(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
; Related .......: _LOCalc_DocRedo, _LOCalc_DocRedoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocRedoGetAllActionTitles(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isRedoPossible()) Then
		Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.UndoManager.getAllRedoActionTitles())
	Else
		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOCalc_DocRedoGetAllActionTitles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocRedoIsPossible
; Description ...: Test whether a Redo action is available to perform for a document.
; Syntax ........: _LOCalc_DocRedoIsPossible(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If the document has a redo action to perform, True is returned, else False.
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
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Document is Read Only or Document has no save location, try SaveAs.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Document Successfully saved.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sFilePath           - a string value. Full path to save the document to, including Filename and extension.
;                  $sFilterName         - [optional] a string value. Default is "". The filter name. Calling "" (blank string), means the filter is chosen automatically based on the file extension. If no extension is present, or if not matched to the list of extensions in this UDF, the .ods extension is used instead, with the filter name of "calc8".
;                  $bOverwrite          - [optional] a boolean value. Default is Null. If True, the existing file will be overwritten.
;                  $sPassword           - [optional] a string value. Default is Null. Sets a password for the document. (Not all file formats can have a Password set). Null or "" (blank string) = No Password.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFilePath not a String.
;				   @Error 1 @Extended 3 Return 0 = $sFilterName not a String.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sPassword not a String.
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

	$sFilePath = _LOCalc_PathConvert($sFilePath, $LOC_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($sFilterName = "") Or ($sFilterName = " ") Then $sFilterName = __LOCalc_FilterNameGet($sFilePath)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	$aProperties[0] = __LOCalc_SetPropertyValue("FilterName", $sFilterName)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOCalc_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	If $sPassword <> Null Then
		If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LOCalc_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	$oDoc.storeAsURL($sFilePath, $aProperties)

	$sSavePath = _LOCalc_PathConvert($sFilePath, $LOC_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOCalc_DocSaveAs

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocToFront
; Description ...: Bring the called document to the front of the other windows.
; Syntax ........: _LOCalc_DocToFront(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Window was successfully brought to the front of the open windows.
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
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
; Name ..........: _LOCalc_DocUndoCurActionTitle
; Description ...: Retrieve the current Undo action Title.
; Syntax ........: _LOCalc_DocUndoCurActionTitle(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
; Related .......: _LOCalc_DocUndo, _LOCalc_DocUndoGetAllActionTitles
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndoCurActionTitle(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isUndoPossible()) Then
		Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.UndoManager.getCurrentUndoActionTitle())
	Else
		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

EndFunc   ;==>_LOCalc_DocUndoCurActionTitle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndoGetAllActionTitles
; Description ...: Retrieve all available Undo action Titles.
; Syntax ........: _LOCalc_DocUndoGetAllActionTitles(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
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
; Related .......: _LOCalc_DocUndo, _LOCalc_DocUndoCurActionTitle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_DocUndoGetAllActionTitles(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oDoc.UndoManager.isUndoPossible()) Then
		Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc.UndoManager.getAllUndoActionTitles())
	Else
		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOCalc_DocUndoGetAllActionTitles

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_DocUndoIsPossible
; Description ...: Test whether a Undo action is available to perform for a document.
; Syntax ........: _LOCalc_DocUndoIsPossible(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If the document has an undo action to perform, True is returned, else False.
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
; Name ..........: _LOCalc_DocVisible
; Description ...: Set or retrieve the current visibility of a document.
; Syntax ........: _LOCalc_DocVisible(ByRef $oDoc[, $bVisible = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the document is visible.
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
; Name ..........: _LOCalc_DocZoom
; Description ...: Modify the zoom value for a document.
; Syntax ........: _LOCalc_DocZoom(ByRef $oDoc[, $iZoomType = Null[, $iZoom = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iZoomType           - [optional] an integer value (0 - 4). Default is Null. The Zoom type, See remarks. See constants $LOC_ZOOMTYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iZoom               - [optional] an integer value (20-600). Default is Null. The zoom percentage. Only valid if Zoom type is set to "By Value"
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iZoomType not an Integer, less than 0 or greater than 4. See constants $LOC_ZOOMTYPE_* as defined in LibreOfficeCalc_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iZoom not an Integer, less than 20 or greater than 600.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 	0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |							1 = Error setting $iZoom
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: Zoom type always has the value of $LOC_ZOOMTYPE_BY_VALUE(3), when using the other zoom types, the value stays the same, but the zoom level is modified. Consequently, I have not added an error check for the  Zoom Type property being correctly set.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
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

	If __LOCalc_VarsAreNull($iZoomType, $iZoom) Then
		__LOCalc_ArrayFill($aiZoom, $oDoc.CurrentController.ZoomType(), $oDoc.CurrentController.ZoomValue())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiZoom)
	EndIf

	If ($iZoomType <> Null) Then
		If Not __LOCalc_IntIsBetween($iZoomType, $LOC_ZOOMTYPE_OPTIMAL, $LOC_ZOOMTYPE_PAGE_WIDTH_EXACT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oDoc.CurrentController.ZoomType = $iZoomType
	EndIf

	If ($iZoom <> Null) Then
		If Not __LOCalc_IntIsBetween($iZoom, 20, 600) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDoc.CurrentController.ZoomValue = $iZoom
		$iError = ($oDoc.CurrentController.ZoomValue() = $iZoom) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOCalc_DocZoom
