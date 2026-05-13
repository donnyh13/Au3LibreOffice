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
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Closing, Saving, etc. L.O. Base documents.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_DocClose
; _LOBase_DocConnect
; _LOBase_DocCreate
; _LOBase_DocDatabaseType
; _LOBase_DocGetName
; _LOBase_DocGetPath
; _LOBase_DocHasPath
; _LOBase_DocIsActive
; _LOBase_DocIsModified
; _LOBase_DocMaximize
; _LOBase_DocMinimize
; _LOBase_DocOpen
; _LOBase_DocSave
; _LOBase_DocSaveAs
; _LOBase_DocSaveCopy
; _LOBase_DocSubComponentsClose
; _LOBase_DocSubComponentsGetList
; _LOBase_DocVisible
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocClose
; Description ...: Close an existing Database Document.
; Syntax ........: _LOBase_DocClose(ByRef $oDoc[, $bSaveChanges = True[, $sSaveName = ""[, $bDeliverOwnership = True]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bSaveChanges        - [optional] Default is True. If True, saves changes if any were made before closing. See remarks.
;                  $sSaveName           - [optional] Default is "". The file name to save the file as, if the file hasn't been saved before. See Remarks.
;                  $bDeliverOwnership   - [optional] Default is True. If True, deliver ownership of the document Object from the script to LibreOffice, recommended is True.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bSaveChanges not a Boolean.
;                  @Error 1 @Extended 3 = $sSaveName not a String.
;                  @Error 1 @Extended 4 = $bDeliverOwnership not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error while creating Filter Name properties.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = $bSaveChanges called with True, and Document hasn't been assigned a Database type yet. Set it using _LOBase_DocDatabaseType.
;                  @Error 3 @Extended 2 = Document hasn't been assigned a Database type yet. Set it using _LOBase_DocDatabaseType.
;                  @Error 3 @Extended 3 = Path Conversion to L.O. URL Failed.
;                  @Error 3 @Extended 4 = Failed to close Document.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Success, Document was successfully closed, and was saved to the returned file Path.
;                  @Error 0 @Extended 2 Return String = Success, Document was successfully closed, document's changes were saved to its existing location.
;                  @Error 0 @Extended 3 Return 1 = Success, Document was successfully closed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bSaveChanges is True and the document hasn't been saved yet, the document is saved to the desktop.
;                  You must set the Database type using _LOBase_DocDatabaseType, before you can save the document that hasn't been saved before.
;                  If $sSaveName is undefined, it is saved as an .odb document to the desktop, named Year-Month-Day_Hour-Minute-Second.odb. $sSaveName may be a name only without an extension, in which case the file will be saved in .odb format, you may also include the extension, such as "Test.odb"
; Related .......: _LOBase_DocOpen, _LOBase_DocConnect, _LOBase_DocCreate, _LOBase_DocSaveAs, _LOBase_DocSave, _LOBase_DocDatabaseType
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocClose(ByRef $oDoc, $bSaveChanges = True, $sSaveName = "", $bDeliverOwnership = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__STR_STRIPLEADING = 1, $__STR_STRIPTRAILING = 2
	Local $sDocPath = "", $sSavePath
	Local $aArgs[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bSaveChanges) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSaveName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDeliverOwnership) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If $bSaveChanges And ($oDoc.DataSource.URL() = "") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($bSaveChanges = True) Then
		If $oDoc.hasLocation() Then
			$oDoc.store()
			$sDocPath = _LO_PathConvert($oDoc.getURL(), $LO_PATHCONV_PCPATH_RETURN)
			$oDoc.Close($bDeliverOwnership)

			If Not __LO_IsObjInvalid($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			$oDoc = Null

			Return SetError($__LO_STATUS_SUCCESS, 2, $sDocPath)

		Else
			If ($oDoc.DataSource.URL() = "") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$sSavePath = @DesktopDir & "\"
			If ($sSaveName = "") Or ($sSaveName = " ") Then
				$sSaveName = @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "-" & @MIN & "-" & @SEC & ".odb"
			EndIf

			$sSaveName = StringStripWS($sSaveName, $__STR_STRIPLEADING + $__STR_STRIPTRAILING)
			If Not StringRegExp($sSaveName, "\Q.odb\E[ ]*$") Then $sSaveName &= ".odb"

			$sSavePath = _LO_PathConvert($sSavePath & $sSaveName, 1)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$aArgs[0] = __LO_SetPropertyValue("FilterName", "StarOffice XML (Base)")
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oDoc.storeAsURL($sSavePath, $aArgs)

			$oDoc.Close($bDeliverOwnership)

			If Not __LO_IsObjInvalid($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			$oDoc = Null

			Return SetError($__LO_STATUS_SUCCESS, 1, _LO_PathConvert($sSavePath, $LO_PATHCONV_PCPATH_RETURN))
		EndIf
	EndIf

	$oDoc.Close($bDeliverOwnership)

	If Not __LO_IsObjInvalid($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$oDoc = Null

	Return SetError($__LO_STATUS_SUCCESS, 3, 1)
EndFunc   ;==>_LOBase_DocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocConnect
; Description ...: Retrieve the Object of an already opened instance of LibreOffice Base.
; Syntax ........: _LOBase_DocConnect([$iMode = $LO_DOC_CONNECT_MODE_CURRENT[, $sSearch = ""[, $bCaseless = False]]])
; Parameters ....: $iMode               - [optional] (0-4) Default is $LO_DOC_CONNECT_MODE_CURRENT. The Connect mode. See Constants, $LO_DOC_CONNECT_MODE_* as defined in LibreOffice_Constants.au3.
;                  $sSearch             - [optional] Default is "". The Name, Title or Path of the Document to search for. See remarks.
;                  $bCaseless           - [optional] Default is False. If True, searches are caseless when using $LO_DOC_CONNECT_MODE_SEARCH_* flags.
; Return values .: Success: Object or Array.
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
;                  @Error 3 @Extended 5 = Current Document not a Base Document.
;                  @Error 3 @Extended 6 = No matches found.
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Success, The Object for the current, or last active Base document is returned.
;                  @Error 0 @Extended 1 Return Object = Success, The Object for the found Document with matching Name, Title or Path.
;                  @Error 0 @Extended ? Return Array = Success, An Array of all open LibreOffice Base Documents. @Extended is set to number of results. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only Base documents are searched or returned using any of the flags.
;                  The value used for $sSearch depends on the flag called in $iMode. It is ignored except for the $LO_DOC_CONNECT_MODE_SEARCH_* flags.
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_TITLE, $sSearch must be the full Title with Office and Component name; e.g: "Test.odb — LibreOffice Base". This will be the same Title AutoIt would match or return from functions like WinGetTitle.
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_NAME, $sSearch must be the Document's full name, without the extension; e.g: "Test".
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_NAME_WITH_EXT, $sSearch must be the Document's name, with the extension; e.g: "Test.odb". If the Document hasn't been saved, just the name will work, e.g., "Untitled 1".
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_PATH, $sSearch must be the full Path of the document (Name and extension included); e.g: "C:\file\Test.odb."
;                  The Connect All option returns a single columned array. ($aArray[0]), each result is stored in a separate row.
;                  -Row 1 contains the Object for that document. e.g. $aArray[0] = $oDoc
;                  -Row 2 contains the Object for the next document. e.g. $aArray[1] = $oDoc2. And so on.
; Related .......: _LOBase_DocOpen, _LOBase_DocClose, _LOBase_DocCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocConnect($iMode = $LO_DOC_CONNECT_MODE_CURRENT, $sSearch = "", $bCaseless = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
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

				If ($iDocType = $LO_DOC_TYPE_BASE) Then
					If (UBound($aoConnectAll) <= $iCount) Then ReDim $aoConnectAll[$iCount + 1]
					$aoConnectAll[$iCount] = $oDoc
					$iCount += 1
				EndIf
				Sleep((IsInt($iCount / $__LOBCONST_SLEEP_DIV) ? (10) : (0)))
			WEnd

			Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoConnectAll)

		Case $LO_DOC_CONNECT_MODE_CURRENT
			$oDoc = $oDesktop.currentComponent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$iDocType = _LO_DocGetType($oDoc)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.
			If ($iDocType <> $LO_DOC_TYPE_BASE) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; Not a Base Doc.

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

				If ($iDocType = $LO_DOC_TYPE_BASE) Then
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
EndFunc   ;==>_LOBase_DocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocCreate
; Description ...: Open a new LibreOffice Base Document.
; Syntax ........: _LOBase_DocCreate([$bForceNew = True[, $bHidden = False[, $bWizard = False]]])
; Parameters ....: $bForceNew           - [optional] Default is True. If True, force opening a new Base Document instead of checking for a usable blank.
;                  $bHidden             - [optional] Default is False. If True opens the new document invisible or changes the existing document to invisible.
;                  $bWizard             - [optional] Default is False. If True, opens the Create a Database Document wizard. See remarks.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $bForceNew not a Boolean.
;                  @Error 1 @Extended 2 = $bHidden not a Boolean.
;                  @Error 1 @Extended 3 = $bWizard not a Boolean.
;                  @Error 1 @Extended 4 = $bWizar and $bHidden both called with True.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Failure Creating Object com.sun.star.ServiceManager.
;                  @Error 2 @Extended 2 = Failure Creating Object com.sun.star.frame.Desktop.
;                  @Error 2 @Extended 3 = Failed to enumerate available documents.
;                  @Error 2 @Extended 4 = Failure Creating New Document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Document Object is still returned. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bHidden
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Successfully connected to an existing Document. Returning Document's Object
;                  @Error 0 @Extended 2 Return Object = Successfully created a new document. Returning Document's Object
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bWizard is True, $bHidden must be False.
;                  If $bWizard is True, the function will not return until the user either cancels or completes the wizard. If the user cancels, an error will result.
;                  You must set the Database type using _LOBase_DocDatabaseType, before you can save the document.
; Related .......: LOBase_DocOpen, LOBase_DocClose, LOBase_DocConnect, _LOBase_DocDatabaseType
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocCreate($bForceNew = True, $bHidden = False, $bWizard = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $aArgs[1]
	Local $iError = 0
	Local $oServiceManager, $oDesktop, $oDoc, $oEnumDoc
	Local $sServiceName = "com.sun.star.sdb.DatabaseDocument"

	If Not IsBool($bForceNew) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bWizard) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $bWizard And $bHidden Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$aArgs[0] = __LO_SetPropertyValue("Hidden", $bHidden)
	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; If not force new, and L.O. pages exist then see if there are any blank Base documents to use.
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

	If $bWizard Then
		$oDoc = $oDesktop.loadComponentFromURL("private:factory/sdatabase?Interactive", "_blank", $iURLFrameCreate, $aArgs)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	Else
		$oDoc = $oDesktop.loadComponentFromURL("private:factory/sdatabase", "_blank", $iURLFrameCreate, $aArgs)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOBase_DocCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocDatabaseType
; Description ...: Set or Retrieve a Base Document's Database Type.
; Syntax ........: _LOBase_DocDatabaseType(ByRef $oDoc[, $sType = "sdbc:embedded:hsqldb"[, $bOverwrite = False]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sType               - [optional] Default is Default. Default is "sdbc:embedded:hsqldb". The Database Type string to set the document to. See remarks.
;                  $bOverwrite          - [optional] Default is False. If True, an existing Database type will be overwritten. See remarks.
; Return values .: Success: 1 or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sType not a String.
;                  @Error 1 @Extended 3 = $bOverwrite not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve current Database type.
;                  @Error 3 @Extended 2 = $bOverwrite is called with False, and Document's Database type is already set.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sType
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. $sType called with Null, returning current Database type as a String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: I have not investigated the various settings for each Database type therefore I have no checks for right or wrong values, if you know the appropriate string to use you can set $sType to the desired setting, but make sure you know what you are doing. "sdbc:embedded:hsqldb" is the default setting for LibreOffice, which creates an embedded HSQLDB Base Document. The Type format is as follows jdbc:subprotocol:subname or sdbc:subprotocol:subname.
;                  I am not knowledgeable enough to know if changing Database types works, or if it is advisable, therefore I made the setting $bOverwrite. If $bOverwrite is False it prevents the user from setting the Database type if one is already set for the document.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword) to get the current Database type.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocDatabaseType(ByRef $oDoc, $sType = Default, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sDataType
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sDataType = $oDoc.DataSource.URL()
	If Not IsString($sDataType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sType) Then

		Return SetError($__LO_STATUS_SUCCESS, 1, $sDataType)

	ElseIf ($sType = Default) Then
		$sType = "sdbc:embedded:hsqldb"
	EndIf

	If Not IsString($sType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sDataType <> "jdbc:") And ($bOverwrite = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDoc.DataSource.URL = $sType
	$iError = ($oDoc.DataSource.URL() = $sType) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_DocDatabaseType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocGetName
; Description ...: Retrieve the document's name.
; Syntax ........: _LOBase_DocGetName(ByRef $oDoc[, $bReturnFull = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bReturnFull         - [optional] Default is False. If True, the full window title is returned, such as is used by AutoIt window related functions.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bReturnFull not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve Document's name.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the document's Name as a String. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bReturnFull is True, the return value will be like: "<Database Doc name>.<extension> — LibreOffice Base" e.g. "Testing.odb — LibreOffice Base".
;                  Else the return value will be like: "<Database Doc name>.<extension>", e.g. "Testing.odb"
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocGetName(ByRef $oDoc, $bReturnFull = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
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
EndFunc   ;==>_LOBase_DocGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocGetPath
; Description ...: Returns a Document's current save path.
; Syntax ........: _LOBase_DocGetPath(ByRef $oDoc[, $bReturnLibreURL = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bReturnLibreURL     - [optional] Default is False. If True, returns a path in LibreOffice URL format, else False returns a regular Windows path.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bReturnLibreURL not a Boolean.
;                  @Error 1 @Extended 3 = Document has no save path.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Error converting LibreOffice URL to Computer path format.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the document's save path as a String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LO_PathConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocGetPath(ByRef $oDoc, $bReturnLibreURL = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
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
EndFunc   ;==>_LOBase_DocGetPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocHasPath
; Description ...: Returns whether a document has been saved to a location already or not.
; Syntax ........: _LOBase_DocHasPath(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether the Document had a path value.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the document has a save location. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocHasPath(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bHasPath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bHasPath = $oDoc.hasLocation()
	If Not IsBool($bHasPath) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bHasPath)
EndFunc   ;==>_LOBase_DocHasPath

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocIsActive
; Description ...: Tests if called document is the active document of other LibreOffice windows.
; Syntax ........: _LOBase_DocIsActive(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether the Document is active.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if document is the currently active LibreOffice window. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This does NOT test if the document is the current active window in Windows, it only tests if the document is the current active document among other LibreOffice documents.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocIsActive(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsActive

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsActive = $oDoc.CurrentController.Frame.isActive()
	If Not IsBool($bIsActive) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsActive)
EndFunc   ;==>_LOBase_DocIsActive

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocIsModified
; Description ...: Test whether the document has been modified since being created or since the last save.
; Syntax ........: _LOBase_DocIsModified(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether the Document has been modified.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the document has been modified since last being saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocIsModified(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsMod

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bIsMod = $oDoc.isModified()
	If Not IsBool($bIsMod) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bIsMod)
EndFunc   ;==>_LOBase_DocIsModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocMaximize
; Description ...: Maximize or restore a document.
; Syntax ........: _LOBase_DocMaximize(ByRef $oDoc[, $bMaximize = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bMaximize           - [optional] Default is Null. If True, document window is maximized, else if False, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bMaximize not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether the Document is Maximized.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bMaximize
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully maximized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMaximize called with Null, returning boolean indicating if Document is currently maximized (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocMaximize(ByRef $oDoc, $bMaximize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
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
EndFunc   ;==>_LOBase_DocMaximize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocMinimize
; Description ...: Minimize or restore a document.
; Syntax ........: _LOBase_DocMinimize(ByRef $oDoc[, $bMinimize = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bMinimize           - [optional] Default is Null. If True, document window is minimized, else if False, document is restored to its previous size and location.
; Return values .: Success: 1 or Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bMinimize not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether the Document is Minimized.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bMinimize
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Document was successfully minimized.
;                  @Error 0 @Extended 1 Return Boolean = Success. $bMinimize called with Null, returning boolean indicating if Document is currently minimized (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocMinimize(ByRef $oDoc, $bMinimize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
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
EndFunc   ;==>_LOBase_DocMinimize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocOpen
; Description ...: Open an existing Database Document.
; Syntax ........: _LOBase_DocOpen($sFilePath[, $bConnectIfOpen = True[, $bHidden = Null[, $bReadOnly = Null[, $sPassword = Null[, $bLoadAsTemplate = Null]]]]])
; Parameters ....: $sFilePath           - Full path and filename of the file to be opened.
;                  $bConnectIfOpen      - [optional] Default is True. If True, connect to the requested document if it is already open. See remarks.
;                  $bHidden             - [optional] Default is Null. If True, opens the document invisibly.
;                  $bReadOnly           - [optional] Default is Null. If True, opens the document as read-only.
;                  $sPassword           - [optional] Default is Null. The password that was used to read-protect the document, if any.
;                  $bLoadAsTemplate     - [optional] Default is Null. If True, opens the document as a Template, i.e. an untitled copy of the specified document is made instead of modifying the original document.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $sFilePath not string, or file not found.
;                  @Error 1 @Extended 2 = Error converting file path to URL path.
;                  @Error 1 @Extended 3 = $bConnectIfOpen not a Boolean.
;                  @Error 1 @Extended 4 = $bHidden not a Boolean.
;                  @Error 1 @Extended 5 = $bReadOnly not a Boolean.
;                  @Error 1 @Extended 6 = $sPassword not a string.
;                  @Error 1 @Extended 7 = $bLoadAsTemplate not a Boolean.
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
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Successfully connected to requested Document without requested parameters. Returning Document's Object.
;                  @Error 0 @Extended 2 Return Object = Successfully opened requested Document with requested parameters. Returning Document's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Any parameters (Hidden, template etc.,) will not be applied when connecting to a document.
; Related .......: _LOBase_DocCreate, _LOBase_DocClose, _LOBase_DocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocOpen($sFilePath, $bConnectIfOpen = True, $bHidden = Null, $bReadOnly = Null, $sPassword = Null, $bLoadAsTemplate = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
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

	If Not __LO_VarsAreNull($bHidden, $bReadOnly, $sPassword, $bLoadAsTemplate) Then
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
	EndIf

	If $bConnectIfOpen Then $oDoc = _LOBase_DocConnect($LO_DOC_CONNECT_MODE_SEARCH_PATH, $sFilePath, True)
	If IsObj($oDoc) Then Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oDoc))

	$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $aoProperties)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oDoc)) : (SetError($__LO_STATUS_SUCCESS, 2, $oDoc))
EndFunc   ;==>_LOBase_DocOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocSave
; Description ...: Save any changes made to a Document.
; Syntax ........: _LOBase_DocSave(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Document is Read Only or Document has no save location, try SaveAs.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Document Successfully saved.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You must set the DataBase type using _LOBase_DocDatabaseType, before you can save the document.
; Related .......: _LOBase_DocSaveAs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocSave(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDoc.hasLocation Or $oDoc.isReadOnly Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oDoc.store()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DocSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocSaveAs
; Description ...: Save a Document with the specified file name to the path specified with any parameters called.
; Syntax ........: _LOBase_DocSaveAs(ByRef $oDoc, $sFilePath[, $bOverwrite = Null[, $sPassword = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sFilePath           - Full path to save the document to, including Filename and extension.
;                  $bOverwrite          - [optional] Default is Null. If True, the existing file will be overwritten.
;                  $sPassword           - [optional] Default is Null. Sets a password for the document. (Not all file formats can have a Password set). Null or "" (blank string) = No Password.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sFilePath not a String.
;                  @Error 1 @Extended 3 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 4 = $sPassword not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error creating FilterName Property
;                  @Error 2 @Extended 2 = Error creating Overwrite Property
;                  @Error 2 @Extended 3 = Error creating Password Property
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Document hasn't been assigned a Database type yet. Set it using _LOBase_DocDatabaseType.
;                  @Error 3 @Extended 2 = Error Converting Path to/from L.O. URL
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Successfully Saved the document. Returning document save path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Alters original save path (if there was one) to the new path.
;                  If ".odb" extension is not present, it is appended to the save path.
;                  You must set the Database type using _LOBase_DocDatabaseType, before you can save the document.
; Related .......: _LOBase_DocSave, _LOBase_DocSaveCopy, _LOBase_DocDatabaseType
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocSaveAs(ByRef $oDoc, $sFilePath, $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__STR_STRIPLEADING = 1, $__STR_STRIPTRAILING = 2
	Local $aProperties[1]
	Local $sSavePath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oDoc.DataSource.URL() = "") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sFilePath = StringStripWS($sFilePath, $__STR_STRIPLEADING + $__STR_STRIPTRAILING)
	If Not StringRegExp($sFilePath, "\Q.odb\E[ ]*$") Then $sFilePath &= ".odb"

	$sFilePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$aProperties[0] = __LO_SetPropertyValue("FilterName", "StarOffice XML (Base)")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	EndIf

	If $sPassword <> Null Then
		If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)
	EndIf

	$oDoc.storeAsURL($sFilePath, $aProperties)

	$sSavePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOBase_DocSaveAs

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocSaveCopy
; Description ...: Save a copy of a Document to the path and file name specified, without modifying the original save location.
; Syntax ........: _LOBase_DocSaveCopy(ByRef $oDoc, $sFilePath[, $bOverwrite = Null[, $sPassword = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $sFilePath           - Full path to save the document to, including Filename and extension. See Remarks.
;                  $bOverwrite          - [optional] Default is Null. If True, file will be overwritten.
;                  $sPassword           - [optional] Default is Null. Password String to set for the document. (Not all file formats can have a Password set). "" (blank string) or Null = No Password.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $sFilePath not a String.
;                  @Error 1 @Extended 3 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 4 = $sPassword not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 = Error creating FilterName Property
;                  @Error 2 @Extended 2 = Error creating Overwrite Property
;                  @Error 2 @Extended 3 = Error creating Password Property
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Document hasn't been assigned a Database type yet. Set it using _LOBase_DocDatabaseType.
;                  @Error 3 @Extended 2 = Error Converting Path to/from L.O. URL
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning save path for exported document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Does not alter the original save path (if there was one), saves a copy of the document to the new path.
; Related .......: _LOBase_DocSave, _LOBase_DocSaveAs, _LOBase_DocDatabaseType
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocSaveCopy(ByRef $oDoc, $sFilePath, $bOverwrite = Null, $sPassword = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aProperties[3]
	Local $sSavePath

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($oDoc.DataSource.URL() = "") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sFilePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$aProperties[0] = __LO_SetPropertyValue("FilterName", "StarOffice XML (Base)")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bOverwrite <> Null) Then
		If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Overwrite", $bOverwrite)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	EndIf

	If ($sPassword <> Null) Then
		If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		ReDim $aProperties[UBound($aProperties) + 1]
		$aProperties[UBound($aProperties) - 1] = __LO_SetPropertyValue("Password", $sPassword)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)
	EndIf

	$oDoc.storeToURL($sFilePath, $aProperties)

	$sSavePath = _LO_PathConvert($sFilePath, $LO_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sSavePath)
EndFunc   ;==>_LOBase_DocSaveCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocSubComponentsClose
; Description ...: Attempt to close all open SubComponent windows.
; Syntax ........: _LOBase_DocSubComponentsClose(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = Document called in $oDoc has not been saved to a location yet.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to attempt to close components.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning a Boolean whether all SubComponents were closed successfully (True), or if some failed to close (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This functions attempts to close all open Sub components (Tables, Queries, Forms or Reports [Except Reports in Viewing mode]). This will fail if any of the following is True for any open components: there are unsaved changes, if a dialog is open or if the user is printing from one of the documents.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocSubComponentsClose(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bClosed

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDoc.hasLocation() Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$bClosed = $oDoc.CurrentController.closeSubComponents()
	If Not IsBool($bClosed) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bClosed)
EndFunc   ;==>_LOBase_DocSubComponentsClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocSubComponentsGetList
; Description ...: Retrieve an Array of currently open SubComponents (Tables, Queries, Forms or Reports [Except Reports in Viewing mode]).
; Syntax ........: _LOBase_DocSubComponentsGetList(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = Document called in $oDoc has not been saved to a location yet.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to retrieve array of SubComponents.
;                  @Error 3 @Extended 2 = Failed to identify SubComponent type.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning a four column Array of currently open SubComponents (Tables, Queries, Forms or Reports [Except Reports in Viewing mode]). @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Return is a 4 column Array.
;                  The First Column is the Component's Object, (Tables, Queries, Forms or Reports [Except Reports in Viewing mode]).
;                  The Second Column is a Constant identifying the type of the component. See Constants, $LOB_SUB_COMP_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  The Third Column is the Component's name, including the path for Forms and Reports, e.g. "frmForm1" or "Folder1/frmForm1".
;                  The Fourth Column is a Boolean indicating whether the Component is in Design mode (True) or Viewing mode (False).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DocSubComponentsGetList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avComponents[0][4], $avSubComponents[0]
	Local $tPropertiesPair

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oDoc.hasLocation() Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$avSubComponents = $oDoc.CurrentController.SubComponents()
	If Not IsArray($avSubComponents) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $avComponents[UBound($avSubComponents)][4]

	For $i = 0 To UBound($avSubComponents) - 1
		$avComponents[$i][0] = $avSubComponents[$i]

		$tPropertiesPair = $oDoc.CurrentController.identifySubComponent($avSubComponents[$i])
		If Not IsObj($tPropertiesPair) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$avComponents[$i][1] = $tPropertiesPair.First() ; Component type. ; See identifySubComponent.
		$avComponents[$i][2] = $tPropertiesPair.Second() ; Component name.

		Switch $avComponents[$i][1]
			Case $LOB_SUB_COMP_TYPE_TABLE
				$avComponents[$i][3] = ($avSubComponents[$i].supportsService("com.sun.star.sdb.TableDesign")) ? (True) : (False) ; If True, Table is in Design mode.

			Case $LOB_SUB_COMP_TYPE_QUERY
				$avComponents[$i][3] = ($avSubComponents[$i].supportsService("com.sun.star.sdb.QueryDesign")) ? (True) : (False) ; If True, Query is in Design mode.

			Case $LOB_SUB_COMP_TYPE_FORM
				$avComponents[$i][3] = ($avSubComponents[$i].isReadOnly) ? (False) : (True) ; If Document is ReadOnly, Form is open in Viewing Mode. Else in Design mode.

			Case $LOB_SUB_COMP_TYPE_REPORT
				$avComponents[$i][3] = True ; The only Reports returned in this are Reports opened in Design Mode.
		EndSwitch
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($avComponents), $avComponents)
EndFunc   ;==>_LOBase_DocSubComponentsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DocVisible
; Description ...: Set or retrieve the current visibility of a document.
; Syntax ........: _LOBase_DocVisible(ByRef $oDoc[, $bVisible = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $bVisible            - [optional] Default is Null. If True, the document is visible.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 = $oDoc not an Object.
;                  @Error 1 @Extended 2 = $bVisible not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 = Failed to query whether the Document is Visible.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
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
Func _LOBase_DocVisible(ByRef $oDoc, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
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
EndFunc   ;==>_LOBase_DocVisible
