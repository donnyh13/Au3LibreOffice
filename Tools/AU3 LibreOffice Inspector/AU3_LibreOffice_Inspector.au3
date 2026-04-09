
; #FUNCTION# ====================================================================================================================
; Name ..........: _AU3_LO_MRI
; Description ...: MRI a LibreOffice Object.
; Syntax ........: _AU3_LO_MRI(ByRef $oObj)
; Parameters ....: $oObj                - an object. Any LibreOffice Object to pass to MRI.
; Return values .: Success: 1.
;                  Failure: 0 and sets @Error to:
;                  |@Error 1 = $oObj not an Object.
;                  |@Error 2 = Failed to create a Document.
;                  |@Error 3 = Failed to retrieve a Marco Object.
;                  |@Error 4 = MRI library doesn't exist. Please install it.
;                  |@Error 5 = Failed to retrieve the MRI Marco Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AU3_LO_MRI(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __AU3_LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc, $oScript, $oLibraryExists
	Local $avParamArray[1], $aDummyArray[0]
	Local $sMacroModuleName = "AU3_LibreOffice_Inspector"

	If Not IsObj($oObj) Then Return SetError(1, 0, 0)

	$oDoc = __AU3_LO_CreateDocument()
	If Not IsObj($oDoc) Then Return SetError(2, 0, 0)

	$oLibraryExists = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard." & $sMacroModuleName & ".LibraryExists?language=Basic&location=document")
	If Not IsObj($oLibraryExists) Then Return SetError(3, 0, __AU3_LO_CloseDoc($oDoc))

	$avParamArray[0] = "MRILib"

	If Not $oLibraryExists.Invoke($avParamArray, $aDummyArray, $aDummyArray) Then Return SetError(4, 0, __AU3_LO_CloseDoc($oDoc))

	$oScript = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard." & $sMacroModuleName & ".MRIThis?language=Basic&location=document")
	If Not IsObj($oScript) Then Return SetError(5, 0, __AU3_LO_CloseDoc($oDoc))

	$avParamArray[0] = $oObj

	$oScript.Invoke($avParamArray, $aDummyArray, $aDummyArray)

	__AU3_LO_CloseDoc($oDoc)

	Return SetError(0, 0, 1)
EndFunc   ;==>_AU3_LO_MRI

; #FUNCTION# ====================================================================================================================
; Name ..........: _AU3_LO_XRAY
; Description ...: XRAY a LibreOffice Object.
; Syntax ........: _AU3_LO_XRAY(ByRef $oObj)
; Parameters ....: $oObj                - an object. Any LibreOffice Object to pass to Xray.
; Return values .: Success: 1.
;                  Failure: 0 and sets @Error to:
;                  |@Error 1 = $oObj not an Object.
;                  |@Error 2 = Failed to create a Document.
;                  |@Error 3 = Failed to retrieve a Marco Object.
;                  |@Error 4 = Xray Tool library doesn't exist. Please install it.
;                  |@Error 5 = Failed to retrieve the Xray Marco Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AU3_LO_XRAY(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __AU3_LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDoc, $oScript, $oLibraryExists
	Local $avParamArray[1], $aDummyArray[0]
	Local $sMacroModuleName = "AU3_LibreOffice_Inspector"

	If Not IsObj($oObj) Then Return SetError(1, 0, 0)

	$oDoc = __AU3_LO_CreateDocument()
	If Not IsObj($oDoc) Then Return SetError(2, 0, 0)

	$oLibraryExists = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard." & $sMacroModuleName & ".LibraryExists?language=Basic&location=document")
	If Not IsObj($oLibraryExists) Then Return SetError(3, 0, __AU3_LO_CloseDoc($oDoc))

	$avParamArray[0] = "XrayTool"

	If Not $oLibraryExists.Invoke($avParamArray, $aDummyArray, $aDummyArray) Then Return SetError(4, 0, __AU3_LO_CloseDoc($oDoc))

	$oScript = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard." & $sMacroModuleName & ".XrayThis?language=Basic&location=document")
	If Not IsObj($oScript) Then Return SetError(5, 0, __AU3_LO_CloseDoc($oDoc))

	$avParamArray[0] = $oObj

	$oScript.Invoke($avParamArray, $aDummyArray, $aDummyArray)

	__AU3_LO_CloseDoc($oDoc)

	Return SetError(0, 0, 1)
EndFunc   ;==>_AU3_LO_XRAY

#Region Internal Functions
; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __AU3_LO_CreateDocument
; Description ...: Create a hidden document to execute a macro in.
; Syntax ........: __AU3_LO_CreateDocument()
; Parameters ....: None
; Return values .: Success: Object.
;                  Failure: 0 and sets @Error and @Extended to:
;                  |@Error 1 @Extended 1 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  |@Error 1 @Extended 2 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  |@Error 1 @Extended 3 = Failed to create a new Document.
;                  |@Error 1 @Extended 4 = Failed to retrieve Standard Library Object.
;                  |@Error 1 @Extended 5 = Failed to insert Macros.
;                  |@Error 2 @Extended ? = Failed to set certain Document flags. Use BitAND for 1 (Hidden flag), 2(MacroExecutionMode flag).
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __AU3_LO_CreateDocument()
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __AU3_LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iURLFrameCreate = 8 ;frame will be created if not found
	Local Const $iMacro_ALWAYS_EXECUTE_NO_WARN = 4
	Local $iError = 0
	Local $oDoc, $oServiceManager, $oDesktop, $oStandardLibrary
	Local $atProperties[2]
	Local $vProperty
	Local $sFileURL = "private:factory/swriter"
	Local $sMacroModuleName = "AU3_LibreOffice_Inspector"
	Local $sMacros = "REM BASIC" & @CRLF & @CRLF & _
			" REM -- LibreOffice Macro for using XRAY." & @CRLF & _
			"Sub XrayThis(vObj)" & @CRLF & _
			"If NOT GlobalScope.BasicLibraries.isLibraryLoaded(""XrayTool"") Then" & @CRLF & _
			"GlobalScope.BasicLibraries.LoadLibrary(""XrayTool"")" & @CRLF & _
			"End If" & @CRLF & _
			"Xray vObj" & @CRLF & _
			"End Sub" & @CRLF & @CRLF & _
			"REM -- LibreOffice Macro for using MRI." & @CRLF & _
			"Sub MRIThis(vObj)" & @CRLF & _
			"If NOT GlobalScope.BasicLibraries.isLibraryLoaded(""MRILib"") Then" & @CRLF & _
			"GlobalScope.BasicLibraries.LoadLibrary(""MRILib"")" & @CRLF & _
			"End If" & @CRLF & _
			"Mri vObj" & @CRLF & _
			"End Sub" & @CRLF & @CRLF & _
			"Function LibraryExists(sLibrary As String)" & @CRLF & _
			"LibraryExists = GlobalScope.BasicLibraries.hasByName(sLibrary)" & @CRLF & _
			"End Function"

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError(1, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError(1, 2, 0)

	$vProperty = __AU3_LO_SetPropertyValue($oServiceManager, "Hidden", True)
	If @error Then $iError = BitOR($iError, 1)
	If Not BitAND($iError, 1) Then $atProperties[0] = $vProperty

	$vProperty = __AU3_LO_SetPropertyValue($oServiceManager, "MacroExecutionMode", $iMacro_ALWAYS_EXECUTE_NO_WARN)
	If @error Then $iError = BitOR($iError, 2)
	If Not BitAND($iError, 2) Then $atProperties[1] = $vProperty

	$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $atProperties)
	If Not IsObj($oDoc) Then Return SetError(1, 3, 0)

	; Retrieving the BasicLibrary.Standard Object fails when using a newly opened document, I found a workaround by updating the following setting.
	$oDoc.BasicLibraries.VBACompatibilityMode = $oDoc.BasicLibraries.VBACompatibilityMode()

	$oStandardLibrary = $oDoc.BasicLibraries.Standard()
	If Not IsObj($oStandardLibrary) Then Return SetError(1, 4, 0)

	If $oStandardLibrary.hasByName($sMacroModuleName) Then $oStandardLibrary.removeByName($sMacroModuleName)

	$oStandardLibrary.insertByName($sMacroModuleName, $sMacros)
	If Not $oStandardLibrary.hasByName($sMacroModuleName) Then Return SetError(1, 5, 0)

	Return ($iError > 0) ? SetError(2, $iError, $oDoc) : SetError(0, 1, $oDoc)
EndFunc   ;==>__AU3_LO_CreateDocument

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __AU3_LO_CloseDoc
; Description ...: Close an opened hidden Document.
; Syntax ........: __AU3_LO_CloseDoc(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A LibreOffice Document Object.
; Return values .: Success: 0.
;                  Failure: 0 and sets @Error to:
;                  |@Error 1 = $oDoc not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __AU3_LO_CloseDoc(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __AU3_LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError(1, 0, 0)

	$oDoc.Close(True)

	Return SetError(0, 0, 0)
EndFunc   ;==>__AU3_LO_CloseDoc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __AU3_LO_CreateStruct
; Description ...: Retrieves a Struct.
; Syntax ........: __AU3_LO_CreateStruct(ByRef $oServiceManager, $sStructName)
; Parameters ....: $oServiceManager     - an object. A LibreOffice Service Manager Object.
;                  $sStructName         - a string value. Name of structure to create.
; Return values .: Success: Structure.
;                  Failure: 0 and sets @Error to:
;                  |@Error 1 = $oServiceManager not an Object.
;                  |@Error 2 = $sStructName not a string.
;                  |@Error 3 = Failed to retrieve requested Structure.
; Author ........: mLipok
; Modified ......: donnyh13 - Added error checking.
; Remarks .......: From WriterDemo.au3 as modified by mLipok from WriterDemo.vbs found in the LibreOffice SDK examples.
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/204665-libreopenoffice-writer/?do=findComment&comment=1471711
; Example .......: No
; ===============================================================================================================================
Func __AU3_LO_CreateStruct(ByRef $oServiceManager, $sStructName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __AU3_LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStruct

	If Not IsObj($oServiceManager) Then Return SetError(1, 0, 0)
	If Not IsString($sStructName) Then Return SetError(2, 0, 0)

	$tStruct = $oServiceManager.Bridge_GetStruct($sStructName)
	If Not IsObj($tStruct) Then Return SetError(3, 0, 0)

	Return SetError(0, 0, $tStruct)
EndFunc   ;==>__AU3_LO_CreateStruct

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __AU3_LO_SetPropertyValue
; Description ...: Creates a property value struct object.
; Syntax ........: __AU3_LO_SetPropertyValue(ByRef $oServiceManager, $sName, $vValue)
; Parameters ....: $oServiceManager     - [in/out] an object. A LibreOffice Service Manager Object.
;                  $sName               - a string value. Property name.
;                  $vValue              - a variant value. Property value.
; Return values .: Success: Object.
;                  Failure: 0 and sets @Error to:
;                  |@Error 1 = $oServiceManager not an Object.
;                  |@Error 2 = $sName not a string.
;                  |@Error 3 = Failed to create a "com.sun.star.beans.PropertyValue" Structure.
; Author ........: Leagnus, GMK
; Modified ......: donnyh13 - added CreateStruct function. Modified variable names.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __AU3_LO_SetPropertyValue(ByRef $oServiceManager, $sName, $vValue)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __AU3_LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tProperties

	If Not IsObj($oServiceManager) Then Return SetError(1, 0, 0)
	If Not IsString($sName) Then Return SetError(2, 0, 0)

	$tProperties = __AU3_LO_CreateStruct($oServiceManager, "com.sun.star.beans.PropertyValue")
	If @error Or Not IsObj($tProperties) Then Return SetError(3, 0, 0)

	$tProperties.Name = $sName
	$tProperties.Value = $vValue

	Return SetError(0, 0, $tProperties)
EndFunc   ;==>__AU3_LO_SetPropertyValue

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __AU3_LO_InternalComErrorHandler
; Description ...: Catch COM Errors and output them to ConsoleWrite.
; Syntax ........: __AU3_LO_InternalComErrorHandler(ByRef $oError)
; Parameters ....: $oError              - [in/out] an object.
; Return values .: None
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __AU3_LO_InternalComErrorHandler(ByRef $oError)

	ConsoleWrite("!	We intercepted a COM Error" & @CRLF & _
			"AU3_LibreOffice_Inspector" & @CRLF & _
			"Number: 0x" & Hex($oError.number, 8) & @CRLF & _
			"Description: " & $oError.windescription & @CRLF & _
			"At line: " & $oError.scriptline & @CRLF & _
			"Source: " & $oError.source & @CRLF & _
			"Error Description: " & $oError.description & @CRLF)

EndFunc   ;==>__AU3_LO_InternalComErrorHandler
#EndRegion
