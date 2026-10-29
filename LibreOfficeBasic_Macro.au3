#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Basic
#include "LibreOfficeBasic_Helper.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Functions for processing, retrieval, and creation of LibreOffice Basic Macros.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBasic_MacroExecute
; _LOBasic_MacroExecuteLinkCreate
; _LOBasic_MacroLibContainerGetObj
; _LOBasic_MacroLibrariesAreModified
; _LOBasic_MacroLibrariesGetNames
; _LOBasic_MacroLibrariesSave
; _LOBasic_MacroLibraryCreate
; _LOBasic_MacroLibraryExists
; _LOBasic_MacroLibraryHasPassword
; _LOBasic_MacroLibraryIsLoaded
; _LOBasic_MacroLibraryLoad
; _LOBasic_MacroLibraryPasswordChange
; _LOBasic_MacroLibraryPasswordIsVerified
; _LOBasic_MacroLibraryPasswordVerify
; _LOBasic_MacroLibraryRemove
; _LOBasic_MacroLibraryRename
; _LOBasic_MacroModuleCreate
; _LOBasic_MacroModuleExists
; _LOBasic_MacroModuleMacros
; _LOBasic_MacroModuleRemove
; _LOBasic_MacroModuleRename
; _LOBasic_MacroModulesGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroExecute
; Description ...: Execute a Macro.
; Syntax ........: _LOBasic_MacroExecute($sMacroLink[, $oDoc = Null[, $avParams = Null]])
; Parameters ....: $sMacroLink          - A Macro execution link returned by a previous _LOBasic_MacroExecuteLinkCreate function.
;                  $oDoc                - [optional] Default is Null. A LibreOffice Document object returned by a previous Document Open, Connect, or Create function. See remarks.
;                  $avParams            - [optional] Default is Null. An array of parameters to pass to the Macro being called. See remarks.
; Return values .: Success: Variable
;                  @Error: 0, @Extended: 0, Return: Variable = Success. Successfully executed Macro. Returning return value of Macro.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $sMacroLink not a String.
;                  @Error: 1, @Extended: 2 = $oDoc not an Object.
;                  @Error: 1, @Extended: 3 = Macro Execution link called in $sMacroLink references a Macro in a Document, and no Document Object was called in $oDoc.
;                  @Error: 1, @Extended: 4 = $avParams not an Array.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create the Service Manager.
;                  @Error: 2, @Extended: 2 = Failed to create "com.sun.star.script.provider.MasterScriptProviderFactory" Object.
;                  @Error: 2, @Extended: 3 = Failed to create a Script provider.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Document Script provider.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Script Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the macro to be called is contained in a Document, $oDoc must be the Document Object containing the Macro.
;                  If the Macro to be called is contained in the Application container, a Document Object doesn't need to be called, but Functions such as ThisComponent will not work.
;                  The each element in the array passed in $avParams must be in the order of the macro parameters.
;                  If the Macro has no parameters, an empty array can be passed, or Null.
;                  If the Macro returns a value, it will be returned from this function.
;                  AutoIt does not support in/out, and out parameters in Macros. In/Out parameters will still work to pass a value to, but it won't be modified.
; Related .......: _LOBasic_MacroExecuteLinkCreate, _LOBasic_MacroModuleMacros
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroExecute($sMacroLink, $oDoc = Null, $avParams = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oMasterScriptProvider, $oScriptProvider, $oScript
	Local $vReturn
	Local $aEmptyArray[0]

	If Not IsString($sMacroLink) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($oDoc <> Null) And Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If StringRegExp($sMacroLink, "\&location=document") And Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($avParams <> Null) And Not IsArray($avParams) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If ($avParams = Null) Then $avParams = $aEmptyArray

	If Not IsObj($oDoc) Then
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		; Found in LO Help file "Calling Python Scripts from Basic".
		$oMasterScriptProvider = $oServiceManager.createInstance("com.sun.star.script.provider.MasterScriptProviderFactory")
		If Not IsObj($oMasterScriptProvider) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oScriptProvider = $oMasterScriptProvider.createScriptProvider("")
		If Not IsObj($oScriptProvider) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Else
		$oScriptProvider = $oDoc.getScriptProvider()
		If Not IsObj($oScriptProvider) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	$oScript = $oScriptProvider.getScript($sMacroLink)
	If Not IsObj($oScript) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$vReturn = $oScript.Invoke($avParams, $aEmptyArray, $aEmptyArray)

	Return SetError($__LO_STATUS_SUCCESS, 0, $vReturn)
EndFunc   ;==>_LOBasic_MacroExecute

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroExecuteLinkCreate
; Description ...: Create a Macro execution link.
; Syntax ........: _LOBasic_MacroExecuteLinkCreate(ByRef $oLibContainer, $sLibrary, $sModule, $sMacro[, $sLanguageType = "Basic"[, $bMacroInDoc = True]])
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name containing the Module.
;                  $sModule             - The Module name containing the Macro.
;                  $sMacro              - The Macro to execute.
;                  $sLanguageType       - [optional] Default is "Basic". The language type of the Macro. Usually this is "Basic".
;                  $bMacroInDoc         - [optional] Default is True. If True, the Macro to be called is contained in a Document, Else it is stored in the LibreOffice Application library.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Success. Returning the formatted Macro Execution link.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  @Error: 1, @Extended: 4 = $sModule not a String.
;                  @Error: 1, @Extended: 5 = Module called in $sModule not found.
;                  @Error: 1, @Extended: 6 = $sMacro not a String.
;                  @Error: 1, @Extended: 7 = Macro called in $sMacro not found in Module.
;                  @Error: 1, @Extended: 8 = $sLanguageType not a String.
;                  @Error: 1, @Extended: 9 = $bMacroInDoc not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Library Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Macro collection string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Macro execution link can be used to execute a macro manually, or to assign a macro to a button, form control, etc.
; Related .......: _LOBasic_MacroExecute, _LOBasic_MacroLibrariesGetNames, _LOBasic_MacroModulesGetNames, _LOBasic_MacroModuleMacros
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroExecuteLinkCreate(ByRef $oLibContainer, $sLibrary, $sModule, $sMacro, $sLanguageType = "Basic", $bMacroInDoc = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oLibrary
	Local $sScript
	Local $sCurrMacros, $sMacroLocation

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oLibrary = $oLibContainer.getByName($sLibrary)
	If Not IsObj($oLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oLibrary.hasByName($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$sCurrMacros = $oLibrary.getByName($sModule)
	If Not IsString($sCurrMacros) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not IsString($sMacro) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not StringRegExp($sCurrMacros, "(?im)^\s*(public|private)?\s*(Sub|Function){1}\s*\Q" & $sMacro & "\E\W") Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not IsString($sLanguageType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If Not IsBool($bMacroInDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	$sMacroLocation = ($bMacroInDoc) ? ("document") : ("application")
	$sScript = "vnd.sun.star.script:" & $sLibrary & "." & $sModule & "." & $sMacro & "?language=" & $sLanguageType & "&location=" & $sMacroLocation

	Return SetError($__LO_STATUS_SUCCESS, 0, $sScript)
EndFunc   ;==>_LOBasic_MacroExecuteLinkCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibContainerGetObj
; Description ...: Retrieve the Library container Object for either a Document, or the LibreOffice application.
; Syntax ........: _LOBasic_MacroLibContainerGetObj([$oDoc = Default])
; Parameters ....: $oDoc                - [optional] Default is Default. A LibreOffice Document object returned by a previous Document Open, Connect, or Create function. See remarks.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning the Library Container Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create the Service Manager.
;                  @Error: 2, @Extended: 2 = Failed to create "com.sun.star.script.ApplicationScriptLibraryContainer" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Library container Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Each Document is a Library container, as well as the LibreOffice Application. If a Document Object is passed, the Library container for the Document is returned. If no Document Object is passed, the LibreOffice Application (Global) is returned, for modifying the Global Libraries and Macros.
; Related .......: _LOBasic_MacroLibraryCreate, _LOBasic_MacroLibrariesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibContainerGetObj($oDoc = Default)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oLibrary, $oServiceManager

	If ($oDoc <> Default) And Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If IsObj($oDoc) Then ; Document Library
		; Retrieving a list of BasicLibraries fails when using a newly opened document, I found a workaround by re-applying the following setting.
		$oDoc.BasicLibraries.VBACompatibilityMode = $oDoc.BasicLibraries.VBACompatibilityMode()

		$oLibrary = $oDoc.BasicLibraries()
		If Not IsObj($oLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else ; Global library
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		; ApplicationScriptLibrary is undocumented. So, this could break in the future.
		$oLibrary = $oServiceManager.createInstance("com.sun.star.script.ApplicationScriptLibraryContainer")
		If Not IsObj($oLibrary) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oLibrary)
EndFunc   ;==>_LOBasic_MacroLibContainerGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibrariesAreModified
; Description ...: See if there are any unsaved changes to the Libraries in the Library Contained.
; Syntax ........: _LOBasic_MacroLibrariesAreModified(ByRef $oLibContainer[, $bModified = Null])
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $bModified           - [optional] Default is Null. If True, sets the Library Container's modified status to True.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Library Container modified status was successfully set.
;                  @Error: 0, @Extended: 1, Return: Boolean = Success. Returning True if the Library Container has been modified since last being saved, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $bModified not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query Library Container whether it has been modified.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bModified
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroLibrariesSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibrariesAreModified(ByRef $oLibContainer, $bModified = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bIsMod
	Local $iError = 0

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bModified) Then
		$bIsMod = $oLibContainer.isModified()
		If Not IsBool($bIsMod) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $bIsMod)
	EndIf

	If Not IsBool($bModified) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oLibContainer.Modified = $bModified
	$iError = ($oLibContainer.isModified() = $bModified) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBasic_MacroLibrariesAreModified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibrariesGetNames
; Description ...: Retrieve an array of Libraries contained in a Library Container.
; Syntax ........: _LOBasic_MacroLibrariesGetNames(ByRef $oLibContainer)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning array of Library names. @Extended is set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve array of Library names.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroLibraryExists, _LOBasic_MacroLibraryCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibrariesGetNames(ByRef $oLibContainer)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asLibraries[0]

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asLibraries = $oLibContainer.getElementNames()
	If Not IsArray($asLibraries) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asLibraries), $asLibraries)
EndFunc   ;==>_LOBasic_MacroLibrariesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibrariesSave
; Description ...: Save any changes made to the Library container.
; Syntax ........: _LOBasic_MacroLibrariesSave(ByRef $oLibContainer)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully saved any changes to the Library container.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to save Library Container changes.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroLibrariesAreModified
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibrariesSave(ByRef $oLibContainer)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oLibContainer.storeLibraries()
	If $oLibContainer.isModified() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBasic_MacroLibrariesSave

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibraryCreate
; Description ...: Create a new Library in a Library Container.
; Syntax ........: _LOBasic_MacroLibraryCreate(ByRef $oLibContainer, $sLibrary)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The unique Library name to create.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully created the Library.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library name called in $sLibrary already exists.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create Library.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroLibraryExists, _LOBasic_MacroLibraryRemove, _LOBasic_MacroModuleCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibraryCreate(ByRef $oLibContainer, $sLibrary)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oLibContainer.createLibrary($sLibrary)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBasic_MacroLibraryCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibraryExists
; Description ...: Check if a Library exists by name.
; Syntax ........: _LOBasic_MacroLibraryExists(ByRef $oLibContainer, $sLibrary)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to look for.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Library exists, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query for Library name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroLibraryCreate, _LOBasic_MacroLibraryRename, _LOBasic_MacroModuleExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibraryExists(ByRef $oLibContainer, $sLibrary)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bExists

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$bExists = $oLibContainer.hasByName($sLibrary)
	If Not IsBool($bExists) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bExists)
EndFunc   ;==>_LOBasic_MacroLibraryExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibraryHasPassword
; Description ...: Check if a Library is password protected.
; Syntax ........: _LOBasic_MacroLibraryHasPassword(ByRef $oLibContainer, $sLibrary)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to check.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Library is password protected, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query Library protection status.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroLibraryPasswordChange, _LOBasic_MacroLibraryPasswordVerify, _LOBasic_MacroLibraryPasswordIsVerified
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibraryHasPassword(ByRef $oLibContainer, $sLibrary)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bHasPassword

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$bHasPassword = $oLibContainer.isLibraryPasswordProtected($sLibrary)
	If Not IsBool($bHasPassword) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bHasPassword)
EndFunc   ;==>_LOBasic_MacroLibraryHasPassword

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibraryIsLoaded
; Description ...: Check if a Library is loaded from storage, and is ready to be called.
; Syntax ........: _LOBasic_MacroLibraryIsLoaded(ByRef $oLibContainer, $sLibrary)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to check if it is loaded.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Library is already loaded, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to query Library load status.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroLibraryLoad
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibraryIsLoaded(ByRef $oLibContainer, $sLibrary)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bLoaded

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$bLoaded = $oLibContainer.isLibraryLoaded($sLibrary)
	If Not IsBool($bLoaded) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bLoaded)
EndFunc   ;==>_LOBasic_MacroLibraryIsLoaded

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibraryLoad
; Description ...: Load a Library from storage, preparing it to be called.
; Syntax ........: _LOBasic_MacroLibraryLoad(ByRef $oLibContainer, $sLibrary)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to load.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Library was successfully loaded
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to load Library.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroLibraryIsLoaded
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibraryLoad(ByRef $oLibContainer, $sLibrary)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oLibContainer.LoadLibrary($sLibrary)
	If Not $oLibContainer.isLibraryLoaded($sLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBasic_MacroLibraryLoad

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibraryPasswordChange
; Description ...: Set, Remove or Modify the password for a library.
; Syntax ........: _LOBasic_MacroLibraryPasswordChange(ByRef $oLibContainer, $sLibrary, $sOldPassword, $sNewPassword)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to modify the password for.
;                  $sOldPassword        - The current password of the Library. See remarks.
;                  $sNewPassword        - The new password to apply to the Library. See remarks.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully changed the Library password.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  @Error: 1, @Extended: 4 = Cannot change the password for built-in libraries, or Read-Only Libraries.
;                  @Error: 1, @Extended: 5 = Cannot change the password for the "Standard" library.
;                  @Error: 1, @Extended: 6 = $sOldPassword not a String.
;                  @Error: 1, @Extended: 7 = $sNewPassword not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to set new Library password.
;                  @Error: 3, @Extended: 2 = Failed to remove Library password.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the library wasn't password protected before, the $sOldPassword parameter has to be an empty string.
;                  If the $sNewPassword parameter is an empty string, and the library was password protected before, the library password protection will be disabled.
;                  If the library is password protected and a wrong Old Password is passed, or if the Library isn't password protected and the Old Password isn't an empty string or the library is read only a com.sun.star.lang.IllegalArgumentException COM error is thrown.
; Related .......: _LOBasic_MacroLibraryHasPassword, _LOBasic_MacroLibraryPasswordVerify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibraryPasswordChange(ByRef $oLibContainer, $sLibrary, $sOldPassword, $sNewPassword)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oLibContainer.isLibraryReadOnly($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($sLibrary = "Standard") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsString($sOldPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsString($sNewPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$oLibContainer.changeLibraryPassword($sLibrary, $sOldPassword, $sNewPassword)
	If ($sNewPassword <> "") And Not $oLibContainer.isLibraryPasswordProtected($sLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($sNewPassword = "") And Not ($oLibContainer.isLibraryPasswordProtected($sLibrary) = False) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBasic_MacroLibraryPasswordChange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibraryPasswordIsVerified
; Description ...: Check if a library has been verified by entering the correct password.
; Syntax ........: _LOBasic_MacroLibraryPasswordIsVerified(ByRef $oLibContainer, $sLibrary)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to check.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Library has been verified by the correct password, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Library is not password protected.
;                  @Error: 3, @Extended: 2 = Failed to query Library password verification status.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroLibraryPasswordVerify, _LOBasic_MacroLibraryHasPassword
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibraryPasswordIsVerified(ByRef $oLibContainer, $sLibrary)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bPasswordIsVerified

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not $oLibContainer.isLibraryPasswordProtected($sLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$bPasswordIsVerified = $oLibContainer.isLibraryPasswordVerified($sLibrary)
	If Not IsBool($bPasswordIsVerified) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bPasswordIsVerified)
EndFunc   ;==>_LOBasic_MacroLibraryPasswordIsVerified

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibraryPasswordVerify
; Description ...: Verify a password protected library.
; Syntax ........: _LOBasic_MacroLibraryPasswordVerify(ByRef $oLibContainer, $sLibrary, $sPassword)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to verify.
;                  $sPassword           - The password of the Library.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Library was successfully verified by the correct password, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  @Error: 1, @Extended: 4 = $sPassword not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Library is not password protected.
;                  @Error: 3, @Extended: 2 = Library is already verified.
;                  @Error: 3, @Extended: 3 = Failed to verify Library.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroLibraryPasswordIsVerified, _LOBasic_MacroLibraryHasPassword, _LOBasic_MacroLibraryPasswordChange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibraryPasswordVerify(ByRef $oLibContainer, $sLibrary, $sPassword)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bPasswordIsVerified

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sPassword) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not $oLibContainer.isLibraryPasswordProtected($sLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oLibContainer.isLibraryPasswordVerified($sLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$bPasswordIsVerified = $oLibContainer.verifyLibraryPassword($sLibrary, $sPassword)
	If Not IsBool($bPasswordIsVerified) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bPasswordIsVerified)
EndFunc   ;==>_LOBasic_MacroLibraryPasswordVerify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibraryRemove
; Description ...: Delete a Library from a Library Container.
; Syntax ........: _LOBasic_MacroLibraryRemove(ByRef $oLibContainer, $sLibrary)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to delete.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully deleted the Library.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  @Error: 1, @Extended: 4 = Cannot delete built-in libraries, or Read-Only Libraries.
;                  @Error: 1, @Extended: 5 = Cannot delete the "Standard" library.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to delete requested Library.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You cannot delete built in Libraries or the "Standard" library. Use caution when deleting libraries programmatically, you could lose your entire library by accident!
; Related .......: _LOBasic_MacroLibraryCreate, _LOBasic_MacroLibrariesGetNames, _LOBasic_MacroModuleRemove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibraryRemove(ByRef $oLibContainer, $sLibrary)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oLibContainer.isLibraryReadOnly($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($sLibrary = "Standard") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oLibContainer.removeLibrary($sLibrary)
	If $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBasic_MacroLibraryRemove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroLibraryRename
; Description ...: Rename a Library.
; Syntax ........: _LOBasic_MacroLibraryRename(ByRef $oLibContainer, $sLibrary, $sName)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to rename.
;                  $sName               - The unique Library name to rename the Library to.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully renamed the Library.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  @Error: 1, @Extended: 4 = Cannot rename built-in libraries, or Read-Only Libraries.
;                  @Error: 1, @Extended: 5 = Cannot rename the "Standard" library.
;                  @Error: 1, @Extended: 6 = $sName not a String.
;                  @Error: 1, @Extended: 7 = Library called in $sName already exists.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to rename Library.
;                  @Error: 3, @Extended: 2 = Failed to delete old Library.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You cannot rename built in Libraries or the "Standard" library.
; Related .......: _LOBasic_MacroLibraryExists, _LOBasic_MacroModuleRename
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroLibraryRename(ByRef $oLibContainer, $sLibrary, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oLibContainer.isLibraryReadOnly($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($sLibrary = "Standard") Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If $oLibContainer.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$oLibContainer.renameLibrary($sLibrary, $sName)
	If Not $oLibContainer.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oLibContainer.hasByName($sLibrary) Then $oLibContainer.removeLibrary($sLibrary) ; When renaming a Library, the old one isn't deleted automatically.
	If $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBasic_MacroLibraryRename

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroModuleCreate
; Description ...: Create a new Module in a Library.
; Syntax ........: _LOBasic_MacroModuleCreate(ByRef $oLibContainer, $sLibrary, $sModule[, $sMacro = Null])
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to create the Module in.
;                  $sModule             - The unique Module name to create.
;                  $sMacro              - [optional] Default is Null. The Macro(s) to insert in the Module. Calling with Null will insert a blank Macro. See remarks.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully created the Module.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  @Error: 1, @Extended: 4 = $sModule not a String.
;                  @Error: 1, @Extended: 5 = Module called in $sModule not found.
;                  @Error: 1, @Extended: 6 = $sMacro not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Cannot create Module in built-in libraries, or Read-Only Libraries.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Library Object.
;                  @Error: 3, @Extended: 3 = Failed to create Module.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You cannot create Modules in built in Libraries.
;                  Modules only contain one string, which contains the whole collection of macros contained in the Module, they are not separated individually.
;                  New lines use the LF (Line Feed) sequence. When setting the Macro(s), all newlines are automatically converted to LF by LibreOffice.
; Related .......: _LOBasic_MacroLibraryCreate, _LOBasic_MacroModuleExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroModuleCreate(ByRef $oLibContainer, $sLibrary, $sModule, $sMacro = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oLibrary
	Local Const $sBlankMacro = "REM  *****  BASIC  *****" & @LF & @LF & "Sub Main" & @LF & @LF & "End Sub"

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If $oLibContainer.isLibraryReadOnly($sLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oLibrary = $oLibContainer.getByName($sLibrary)
	If Not IsObj($oLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If $oLibrary.hasByName($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($sMacro <> Null) And Not IsString($sMacro) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If ($sMacro = Null) Then $sMacro = $sBlankMacro

	$oLibrary.insertByName($sModule, $sMacro)
	If Not $oLibrary.hasByName($sModule) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBasic_MacroModuleCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroModuleExists
; Description ...: Check if a Module exists in a certain Library.
; Syntax ........: _LOBasic_MacroModuleExists(ByRef $oLibContainer, $sLibrary, $sModule)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to search for the Module in.
;                  $sModule             - The Module name to search for.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Module exists in the Library, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  @Error: 1, @Extended: 4 = $sModule not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Library Object.
;                  @Error: 3, @Extended: 2 = Failed to query for Module name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroModulesGetNames, _LOBasic_MacroModuleCreate, _LOBasic_MacroLibraryExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroModuleExists(ByRef $oLibContainer, $sLibrary, $sModule)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oLibrary
	Local $bExists

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oLibrary = $oLibContainer.getByName($sLibrary)
	If Not IsObj($oLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$bExists = $oLibrary.hasByName($sModule)
	If Not IsBool($bExists) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bExists)
EndFunc   ;==>_LOBasic_MacroModuleExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroModuleMacros
; Description ...: Set or Retrieve the Macros contained in a Module.
; Syntax ........: _LOBasic_MacroModuleMacros(ByRef $oLibContainer, $sLibrary, $sModule[, $sMacros = Null])
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name containing the Module.
;                  $sModule             - The Module name to Set or Retrieve the Macros for.
;                  $sMacros             - [optional] Default is Null. The Macro string to the the Module to. See remarks.
; Return values .: Success: 1 or String.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. All optional parameters were called with Null, returning current Macro collection as a String.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  @Error: 1, @Extended: 4 = $sModule not a String.
;                  @Error: 1, @Extended: 5 = Module called in $sModule not found.
;                  @Error: 1, @Extended: 6 = $sMacros not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Library Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Macro collection string.
;                  @Error: 3, @Extended: 3 = Library is read-only. Can't set Macro string.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sMacros
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Modules only contain one string, which contains the whole collection of macros contained in the Module, they are not separated individually.
;                  New lines use the LF (Line Feed) sequence. When setting the Macro(s), all newlines are automatically converted to LF by LibreOffice.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOBasic_MacroModulesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroModuleMacros(ByRef $oLibContainer, $sLibrary, $sModule, $sMacros = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__STR_STRIPALL = 8
	Local $iError = 0
	Local $oLibrary
	Local $sCurrMacros

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oLibrary = $oLibContainer.getByName($sLibrary)
	If Not IsObj($oLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oLibrary.hasByName($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If __LO_VarsAreNull($sMacros) Then
		$sCurrMacros = $oLibrary.getByName($sModule)
		If Not IsString($sCurrMacros) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrMacros)
	EndIf

	If $oLibContainer.isLibraryReadOnly($sLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If Not IsString($sMacros) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oLibrary.replaceByName($sModule, $sMacros)
	$iError = (StringStripWS($oLibrary.getByName($sModule), $__STR_STRIPALL) = StringStripWS($sMacros, $__STR_STRIPALL)) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBasic_MacroModuleMacros

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroModuleRemove
; Description ...: Delete a Module from a Library.
; Syntax ........: _LOBasic_MacroModuleRemove(ByRef $oLibContainer, $sLibrary, $sModule)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to delete the Module from.
;                  $sModule             - The Module name to delete.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully deleted the Module.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  @Error: 1, @Extended: 4 = Cannot delete built-in libraries, or Read-Only Libraries.
;                  @Error: 1, @Extended: 5 = $sModule not a String.
;                  @Error: 1, @Extended: 6 = Module called in $sModule not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Library Object.
;                  @Error: 3, @Extended: 2 = Failed to delete requested Module.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You cannot delete Modules from built in Libraries.
; Related .......: _LOBasic_MacroModulesGetNames, _LOBasic_MacroModuleCreate, _LOBasic_MacroLibraryRemove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroModuleRemove(ByRef $oLibContainer, $sLibrary, $sModule)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oLibrary

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oLibContainer.isLibraryReadOnly($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oLibrary = $oLibContainer.getByName($sLibrary)
	If Not IsObj($oLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oLibrary.hasByName($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oLibrary.removeByName($sModule)
	If $oLibrary.hasByName($sModule) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBasic_MacroModuleRemove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroModuleRename
; Description ...: Rename a Module.
; Syntax ........: _LOBasic_MacroModuleRename(ByRef $oLibContainer, $sLibrary, $sModule, $sName)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to rename the Module in.
;                  $sModule             - The Module name to rename.
;                  $sName               - The unique Module name to set the Module to.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully renamed the Module.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  @Error: 1, @Extended: 4 = $sModule not a String.
;                  @Error: 1, @Extended: 5 = Module called in $sModule not found.
;                  @Error: 1, @Extended: 6 = $sName not a String.
;                  @Error: 1, @Extended: 7 = Module called in $sName already exists.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Cannot rename Modules in built-in libraries, or Read-Only Libraries.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Library Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve current Macros in Module.
;                  @Error: 3, @Extended: 4 = Failed to insert new Module.
;                  @Error: 3, @Extended: 5 = Failed to remove old Module.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You cannot rename Modules in built in Libraries.
; Related .......: _LOBasic_MacroModuleExists, _LOBasic_MacroLibraryRename
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroModuleRename(ByRef $oLibContainer, $sLibrary, $sModule, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oLibrary
	Local $sMacro

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If $oLibContainer.isLibraryReadOnly($sLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oLibrary = $oLibContainer.getByName($sLibrary)
	If Not IsObj($oLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oLibrary.hasByName($sModule) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If $oLibrary.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$sMacro = $oLibrary.getByName($sModule)
	If Not IsString($sMacro) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oLibrary.insertByName($sName, $sMacro)
	If Not $oLibrary.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$oLibrary.removeByName($sModule)
	If $oLibrary.hasByName($sModule) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBasic_MacroModuleRename

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_MacroModulesGetNames
; Description ...: Retrieve an array of Module names contained in a Library
; Syntax ........: _LOBasic_MacroModulesGetNames(ByRef $oLibContainer, $sLibrary)
; Parameters ....: $oLibContainer       - A Library Container object returned by a previous _LOBasic_MacroLibContainerGetObj function.
;                  $sLibrary            - The Library name to retrieve the array of Module names for.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning array of Module names. @Extended is set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oLibContainer not an Object.
;                  @Error: 1, @Extended: 2 = $sLibrary not a String.
;                  @Error: 1, @Extended: 3 = Library called in $sLibrary not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Library Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve array of Module names.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBasic_MacroModuleExists, _LOBasic_MacroModuleRemove, _LOBasic_MacroLibrariesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_MacroModulesGetNames(ByRef $oLibContainer, $sLibrary)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBasic_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asModules[0]
	Local $oLibrary

	If Not IsObj($oLibContainer) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oLibContainer.hasByName($sLibrary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oLibrary = $oLibContainer.getByName($sLibrary)
	If Not IsObj($oLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asModules = $oLibrary.getElementNames()
	If Not IsArray($asModules) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asModules), $asModules)
EndFunc   ;==>_LOBasic_MacroModulesGetNames
