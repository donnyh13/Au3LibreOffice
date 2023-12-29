#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Constants.au3"
#include "LibreOfficeCalc_Helper.au3"



; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Various functions for internal data processing, data retrieval, retrieving and applying settings for LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __LOCalc_AddTo1DArray
; __LOCalc_ArrayFill
; __LOCalc_CreateStruct
; __LOCalc_FilterNameGet
; __LOCalc_InternalComErrorHandler
; __LOCalc_IntIsBetween
; __LOCalc_NumIsBetween
; __LOCalc_SetPropertyValue
; __LOCalc_UnitConvert
; __LOCalc_VarsAreDefault
; __LOCalc_VarsAreNull
; __LOCalc_VersionCheck
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_AddTo1DArray
; Description ...: Add data to a 1 Dimensional array.
; Syntax ........: __LOCalc_AddTo1DArray(ByRef $aArray, $vData[, $bCountInFirst = False])
; Parameters ....: $aArray              - [in/out] an array of unknowns. The Array to directly add data to. Array will be directly modified.
;                  $vData               - a variant value. The Data to add to the Array.
;                  $bCountInFirst       - [optional] a boolean value. Default is False. If True the first element of the array is a count of contained elements.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $aArray not an Array
;				   @Error 1 @Extended 2 Return 0 = $bCountinFirst not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $aArray contains too many columns.
;				   @Error 1 @Extended 4 Return 0 = $aArray[0] contains non integer data or is not empty, and $bCountInFirst is set to True.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Array item was successfully added.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_AddTo1DArray(ByRef $aArray, $vData, $bCountInFirst = False)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($aArray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bCountInFirst) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If UBound($aArray, $UBOUND_COLUMNS) > 1 Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Too many columns

	If $bCountInFirst And (UBound($aArray) = 0) Then
		ReDim $aArray[1]
		$aArray[0] = 0
	EndIf

	If $bCountInFirst And (($aArray[0] <> "") And Not IsInt($aArray[0])) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	ReDim $aArray[UBound($aArray) + 1]
	$aArray[UBound($aArray) - 1] = $vData
	If $bCountInFirst Then $aArray[0] += 1
	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOCalc_AddTo1DArray

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_ArrayFill
; Description ...: Fill an Array with data.
; Syntax ........: __LOCalc_ArrayFill(ByRef $aArrayToFill[, $vVar1 = Null[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null[, $vVar9 = Null[, $vVar10 = Null[, $vVar11 = Null[, $vVar12 = Null[, $vVar13 = Null[, $vVar14 = Null[, $vVar15 = Null[, $vVar16 = Null[, $vVar17 = Null[, $vVar18 = Null]]]]]]]]]]]]]]]]]])
; Parameters ....: $aArrayToFill        - [in/out] an array of unknowns. The Array to Fill. Array will be directly modified.
;                  $vVar1               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar2               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar3               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar4               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar5               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar6               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar7               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar8               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar9               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar10              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar11              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar12              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar13              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar14              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar15              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar16              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar17              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar18              - [optional] a variant value. Default is Null. The Data to add to the Array.
; Return values .: None
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call only how many you parameters you need to add to the Array. Automatically resizes the Array if it is the incorrect size.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_ArrayFill(ByRef $aArrayToFill, $vVar1 = Null, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null, $vVar9 = Null, $vVar10 = Null, $vVar11 = Null, $vVar12 = Null, $vVar13 = Null, _
		$vVar14 = Null, $vVar15 = Null, $vVar16 = Null, $vVar17 = Null, $vVar18 = Null)
	#forceref $vVar1, $vVar2, $vVar3, $vVar4, $vVar5, $vVar6, $vVar7, $vVar8, $vVar9, $vVar10, $vVar11, $vVar12, $vVar13, $vVar14, $vVar15, $vVar16, $vVar17, $vVar18

	If UBound($aArrayToFill) < (@NumParams - 1) Then ReDim $aArrayToFill[@NumParams - 1]
	For $i = 0 To @NumParams - 2
		$aArrayToFill[$i] = Eval("vVar" & $i + 1)
	Next
EndFunc   ;==>__LOCalc_ArrayFill

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CreateStruct
; Description ...: Creates a Struct.
; Syntax ........: __LOCalc_CreateStruct($sStructName)
; Parameters ....: $sStructName	- a string value. Name of structure to create.
; Return values .:Success: Structure.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sStructName not a string
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object
;				   @Error 2 @Extended 2 Return 0 = Error creating requested structure.
;				   --Success--
;				   @Error 0 @Extended 0 Return Structure = Success. Property Structure Returned
; Author ........: mLipok
; Modified ......: donnyh13 - Added error checking.
; Remarks .......: From WriterDemo.au3 as modified by mLipok from WriterDemo.vbs found in the LibreOffice SDK examples.
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/204665-libreopenoffice-writer/?do=findComment&comment=1471711
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CreateStruct($sStructName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $tStruct

	If Not IsString($sStructName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$tStruct = $oServiceManager.Bridge_GetStruct($sStructName)
	If Not IsObj($tStruct) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $tStruct)
EndFunc   ;==>__LOCalc_CreateStruct

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_FilterNameGet
; Description ...: Retrieves the correct L.O. Filtername for use in SaveAs and Export.
; Syntax ........: __LOCalc_FilterNameGet(ByRef $sDocSavePath[, $bIncludeExportFilters = False])
; Parameters ....: $sDocSavePath           - [in/out] a string value. Full path with extension.
;                  $bIncludeExportFilters  - [optional] a boolean value. Default is False. If True, includes the FilterNames that can be used to Export only, in the search.
; Return values .:Success: String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sDocSavePath is not a string.
;				   @Error 1 @Extended 2 Return 0 = $bIncludeExportFilters not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sDocSavePath is not a correct path or URL.
;				   --Success--
;				   @Error 0 @Extended 1 Return String = Success. Returns required filtername from "SaveAs" FilterNames.
;				   @Error 0 @Extended 2 Return String = Success. Returns required filtername from "Export" FilterNames.
;				   @Error 0 @Extended 3 Return String = FilterName not found for given file extension, defaulting to .ods file format and updating save path accordingly.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Searches a predefined list of extensions stored in an array. Not all FilterNames are listed.
;				   For finding your own FilterNames, see convertfilters.html found in
;						L.O. Install Folder: LibreOffice\help\en-US\text\shared\guide
;				   Or See: "OOME_3_0",	"OpenOffice.org Macros Explained OOME Third Edition" by Andrew D. Pitonyak, which has a handy Macro for
;						listing all FilterNames, found on page 284 of the above book in the ODT format.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_FilterNameGet(ByRef $sDocSavePath, $bIncludeExportFilters = False)
	Local $iLength, $iSlashLocation, $iDotLocation
	Local Const $STR_NOCASESENSE = 0, $STR_STRIPALL = 8
	Local $sFileExtension, $sFilterName
	Local $msSaveAsFilters[], $msExportFilters[]

	If Not IsString($sDocSavePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bIncludeExportFilters) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$iLength = StringLen($sDocSavePath)

	$msSaveAsFilters[".csv"] = "Text - txt - csv (StarCalc)"
	$msSaveAsFilters[".dbf"] = "dBase"
	$msSaveAsFilters[".dif"] = "DIF"
	$msSaveAsFilters[".et"] = "MS Excel 97"
	$msSaveAsFilters[".ett"] = "MS Excel 97 Vorlage/Template"
	$msSaveAsFilters[".fods"] = "OpenDocument Spreadsheet Flat XML"
	$msSaveAsFilters[".htm"] = "HTML (StarCalc)"
	$msSaveAsFilters[".html"] = "HTML (StarCalc)"
	$msSaveAsFilters[".ods"] = "calc8"
	$msSaveAsFilters[".ots"] = "calc8_template"
	$msSaveAsFilters[".slk"] = "SYLK"
	$msSaveAsFilters[".sylk"] = "SYLK"
	$msSaveAsFilters[".tab"] = "Text - txt - csv (StarCalc)"
	$msSaveAsFilters[".tsv"] = "Text - txt - csv (StarCalc)"
	$msSaveAsFilters[".txt"] = "Text - txt - csv (StarCalc)"
	$msSaveAsFilters[".uof"] = "UOF spreadsheet"
	$msSaveAsFilters[".uos"] = "UOF spreadsheet"
	$msSaveAsFilters[".xhtml"] = "HTML (StarCalc)"
	$msSaveAsFilters[".xlc"] = "MS Excel 97"
	$msSaveAsFilters[".xlk"] = "MS Excel 97"
	$msSaveAsFilters[".xlm"] = "MS Excel 97"
	$msSaveAsFilters[".xls"] = "MS Excel 97"
	$msSaveAsFilters[".xlsm"] = "Calc MS Excel 2007 VBA XML"
	$msSaveAsFilters[".xlsx"] = "Calc MS Excel 2007 XML"
	$msSaveAsFilters[".xlt"] = "MS Excel 97 Vorlage/Template"
	$msSaveAsFilters[".xltm"] = "Calc MS Excel 2007 XML Template"
	$msSaveAsFilters[".xltx"] = "Calc MS Excel 2007 XML Template"
	$msSaveAsFilters[".xlw"] = "MS Excel 97"
	$msSaveAsFilters[".xml"] = "OpenDocument Spreadsheet Flat XML"

	If $bIncludeExportFilters Then
		$msExportFilters[".jfif"] = "calc_jpg_Export"
		$msExportFilters[".jif"] = "calc_jpg_Export"
		$msExportFilters[".jpe"] = "calc_jpg_Export"
		$msExportFilters[".jpeg"] = "calc_jpg_Export"
		$msExportFilters[".jpg"] = "calc_jpg_Export"
		$msExportFilters[".pdf"] = "calc_pdf_Export"
		$msExportFilters[".png"] = "calc_png_Export"
	EndIf

	If StringInStr($sDocSavePath, "file:///") Then ;  If L.O. URl Then
		$iSlashLocation = StringInStr($sDocSavePath, "/", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, ($iLength - $iDotLocation) + 1)
	ElseIf StringInStr($sDocSavePath, "\") Then ;  Else if PC Path Then
		$iSlashLocation = StringInStr($sDocSavePath, "\", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, $iLength - $iDotLocation + 1)
	Else
		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	If $sFileExtension = $sDocSavePath Then ;  If no file extension identified, append .ods extension and return.
		$sDocSavePath = $sDocSavePath & ".ods"
		Return SetError($__LO_STATUS_SUCCESS, 3, "calc8")
	Else
		$sFileExtension = StringLower(StringStripWS($sFileExtension, $STR_STRIPALL))
	EndIf

	$sFilterName = $msSaveAsFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sFilterName)

	If $bIncludeExportFilters Then $sFilterName = $msExportFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 2, $sFilterName)

	$sDocSavePath = StringReplace($sDocSavePath, $sFileExtension, ".ods") ; If No results, replace with ODS extension.

	Return SetError($__LO_STATUS_SUCCESS, 3, "calc8")
EndFunc   ;==>__LOCalc_FilterNameGet

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LOCalc_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _LOCalc_ComError_UserFunction(Default)
	Local $vUserFunction, $avUserParams[2] = ["CallArgArray", $oComError]

	If IsArray($avUserFunction) Then
		$vUserFunction = $avUserFunction[0]
		ReDim $avUserParams[UBound($avUserFunction) + 1]
		For $i = 1 To UBound($avUserFunction) - 1
			$avUserParams[$i + 1] = $avUserFunction[$i]
		Next
	Else
		$vUserFunction = $avUserFunction
	EndIf
	If IsFunc($vUserFunction) Then
		Switch $vUserFunction
			Case ConsoleWrite
				ConsoleWrite("!--COM Error-Begin--" & @CRLF & _
						"Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline & @CRLF & _
						"!--COM-Error-End--" & @CRLF)
			Case MsgBox
				MsgBox(0, "COM Error", "Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline)
			Case Else
				Call($vUserFunction, $avUserParams)
		EndSwitch
	EndIf
EndFunc   ;==>__LOCalc_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_IntIsBetween
; Description ...: Test whether an input is an Integer and is between two Numbers.
; Syntax ........: __LOCalc_IntIsBetween($iTest, $nMin, $nMax[, $snNot = ""[, $snIncl = Default]])
; Parameters ....: $iTest               - an integer value. The Value to test.
;                  $nMin                - a general number value. The minimum $iTest can be.
;                  $nMax                - a general number value. The maximum $iTest can be.
;                  $snNot               - [optional] a string value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers inside the min/max range that are not allowed.
;                  $snIncl              - [optional] a string value. Default is Default. Can be a single number, or a String of numbers separated by ":". Defines numbers Outside the min/max range that are allowed.
; Return values .: Success: Boolean
;				   Failure: False
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If the input is between Min and Max or is an allowed number, and not one of the disallowed numbers, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_IntIsBetween($iTest, $nMin, $nMax, $snNot = "", $snIncl = Default)
	Local $bMatch = False
	Local $anNot, $anIncl

	If Not IsInt($iTest) Then Return False
	If (@NumParams = 3) Then Return (($iTest < $nMin) Or ($iTest > $nMax)) ? (False) : (True)

	If ($snNot <> "") Then
		If IsString($snNot) And StringInStr($snNot, ":") Then
			$anNot = StringSplit($snNot, ":")
			For $i = 1 To $anNot[0]
				If ($anNot[$i] = $iTest) Then Return False
			Next
		Else
			If ($iTest = $snNot) Then Return False
		EndIf
	EndIf

	If (($iTest >= $nMin) And ($iTest <= $nMax)) Then Return True

	If IsString($snIncl) And StringInStr($snIncl, ":") Then
		$anIncl = StringSplit($snIncl, ":")
		For $j = 1 To $anIncl[0]
			$bMatch = ($anIncl[$j] = $iTest) ? (True) : (False)
			If $bMatch Then ExitLoop
		Next
	ElseIf IsNumber($snIncl) Then
		$bMatch = ($iTest = $snIncl) ? (True) : (False)
	EndIf

	Return $bMatch
EndFunc   ;==>__LOCalc_IntIsBetween

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_NumIsBetween
; Description ...: Test whether an input is a Number and is between two Numbers.
; Syntax ........: __LOCalc_NumIsBetween($nTest, $nMin, $nMax[, $snNot = ""[, $snIncl = Default]])
; Parameters ....: $nTest               - a general number value. The Value to test.
;                  $nMin                - a general number value. The minimum $iTest can be.
;                  $nMax                - a general number value. The maximum $iTest can be.
;                  $snNot               - [optional] a string value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers inside the min/max range that are not allowed.
;                  $snIncl              - [optional] a string value. Default is Default. Can be a single number, or a String of numbers separated by ":". Defines numbers Outside the min/max range that are allowed.
; Return values .: Success: Boolean
;				   Failure: False
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If the input is between Min and Max or is an allowed number, and not one of the disallowed numbers, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_NumIsBetween($nTest, $nMin, $nMax, $snNot = "", $snIncl = Default)
	Local $bMatch = False
	Local $anNot, $anIncl

	If Not IsNumber($nTest) Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
	If (@NumParams = 3) Then Return (($nTest < $nMin) Or ($nTest > $nMax)) ? (SetError($__LO_STATUS_SUCCESS, 0, False)) : (SetError($__LO_STATUS_SUCCESS, 0, True))

	If ($snNot <> "") Then
		If IsString($snNot) And StringInStr($snNot, ":") Then
			$anNot = StringSplit($snNot, ":")
			For $i = 1 To $anNot[0]
				If ($anNot[$i] = $nTest) Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
			Next
		Else
			If ($nTest = $snNot) Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		EndIf
	EndIf

	If (($nTest >= $nMin) And ($nTest <= $nMax)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	If IsString($snIncl) And StringInStr($snIncl, ":") Then
		$anIncl = StringSplit($snIncl, ":")
		For $j = 1 To $anIncl[0]
			$bMatch = ($anIncl[$j] = $nTest) ? (True) : (False)
			If $bMatch Then ExitLoop
		Next
	ElseIf IsNumber($snIncl) Then
		$bMatch = ($nTest = $snIncl) ? (True) : (False)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $bMatch)
EndFunc   ;==>__LOCalc_NumIsBetween

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_SetPropertyValue
; Description ...: Creates a property value struct object.
; Syntax ........: __LOCalc_SetPropertyValue($sName, $vValue)
; Parameters ....: $sName               - a string value. Property name.
;                  $vValue              - a variant value. Property value.
; Return values .:Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sName not a string
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create Properties Structure.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Property Object Returned
; Author ........: Leagnus, GMK
; Modified ......: donnyh13 - added CreateStruct function. Modified variable names.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_SetPropertyValue($sName, $vValue)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tProperties

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	$tProperties = __LOCalc_CreateStruct("com.sun.star.beans.PropertyValue")
	If @error Or Not IsObj($tProperties) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$tProperties.Name = $sName
	$tProperties.Value = $vValue

	Return SetError($__LO_STATUS_SUCCESS, 0, $tProperties)
EndFunc   ;==>__LOCalc_SetPropertyValue

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_UnitConvert
; Description ...: For converting measurement units.
; Syntax ........: __LOCalc_UnitConvert($nValue, $sReturnType)
; Parameters ....: $nValue              - a general number value. The Number to be converted.
;                  $iReturnType         - a Integer value. Determines conversion type. See Constants, $__LOCONST_CONVERT_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .:Success: Integer or Number.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $nValue is not a Number.
;				   @Error 1 @Extended 2 Return 0 = $iReturnType is not a Integer.
;				   @Error 1 @Extended 3 Return 0 = $iReturnType does not match constants, See Constants, $__LOCONST_CONVERT_* as defined in LibreOfficeCalc_Constants.au3.
;				   --Success--
;				   @Error 0 @Extended 1 Return Number = Returns Number converted from TWIPS to Centimeters.
;				   @Error 0 @Extended 2 Return Number = Returns Number converted from TWIPS to Inches.
;				   @Error 0 @Extended 3 Return Integer = Returns Number converted from Millimeters to uM (Micrometers).
;				   @Error 0 @Extended 4 Return Number = Returns Number converted from Micrometers to MM
;				   @Error 0 @Extended 5 Return Integer = Returns Number converted from Centimeters To uM
;				   @Error 0 @Extended 6 Return Number = Returns Number converted from um (Micrometers) To CM
;				   @Error 0 @Extended 7 Return Integer = Returns Number converted from Inches to uM(Micrometers).
;				   @Error 0 @Extended 8 Return Number = Returns Number converted from uM(Micrometers) to Inches.
;				   @Error 0 @Extended 9 Return Integer = Returns Number converted from TWIPS to uM(Micrometers).
;				   @Error 0 @Extended 10 Return Integer = Returns Number converted from Point to uM(Micrometers).
;				   @Error 0 @Extended 11 Return Number = Returns Number converted from uM(Micrometers) to Point.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_ConvertFromMicrometer, _LOCalc_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_UnitConvert($nValue, $iReturnType)
	Local $iUM, $iMM, $iCM, $iInch

	If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iReturnType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch $iReturnType

		Case $__LOCONST_CONVERT_TWIPS_CM ;TWIPS TO CM
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = ($nValue / 20 / 72)
			; 1 Inch = 2.54 CM
			$iCM = Round(Round($iInch * 2.54, 3), 2)
			Return SetError($__LO_STATUS_SUCCESS, 1, Number($iCM))

		Case $__LOCONST_CONVERT_TWIPS_INCH ;TWIPS to Inch
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = ($nValue / 20 / 72)
			$iInch = Round(Round($iInch, 3), 2)
			Return SetError($__LO_STATUS_SUCCESS, 2, Number($iInch))

		Case $__LOCONST_CONVERT_MM_UM ;Millimeter to Micrometer
			$iUM = ($nValue * 100)
			$iUM = Round(Round($iUM, 1))
			Return SetError($__LO_STATUS_SUCCESS, 3, Number($iUM))

		Case $__LOCONST_CONVERT_UM_MM ;Micrometer to Millimeter
			$iMM = ($nValue / 100)
			$iMM = Round(Round($iMM, 3), 2)
			Return SetError($__LO_STATUS_SUCCESS, 4, Number($iMM))

		Case $__LOCONST_CONVERT_CM_UM ;Centimeter to Micrometer
			$iUM = ($nValue * 1000)
			$iUM = Round(Round($iUM, 1))
			Return SetError($__LO_STATUS_SUCCESS, 5, Int($iUM))

		Case $__LOCONST_CONVERT_UM_CM ;Micrometer to Centimeter
			$iCM = ($nValue / 1000)
			$iCM = Round(Round($iCM, 3), 2)
			Return SetError($__LO_STATUS_SUCCESS, 6, Number($iCM))

		Case $__LOCONST_CONVERT_INCH_UM ;Inch to Micrometer
			; 1 Inch - 2.54 Cm; Micrometer = 1/1000 CM
			$iUM = ($nValue * 2.54) * 1000 ; + .0055
			$iUM = Round(Round($iUM, 1))
			Return SetError($__LO_STATUS_SUCCESS, 7, Int($iUM))

		Case $__LOCONST_CONVERT_UM_INCH ;Micrometer to Inch
			; 1 Inch - 2.54 Cm; Micrometer = 1/1000 CM
			$iInch = ($nValue / 1000) / 2.54 ; + .0055
			$iInch = Round(Round($iInch, 3), 2)
			Return SetError($__LO_STATUS_SUCCESS, 8, $iInch)

		Case $__LOCONST_CONVERT_TWIPS_UM ;TWIPS to Micrometer
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = (($nValue / 20) / 72)
			$iInch = Round(Round($iInch, 3), 2)
			; 1 Inch - 25.4 MM; Micrometer = 1/100 MM
			$iUM = Round($iInch * 25.4 * 100)
			Return SetError($__LO_STATUS_SUCCESS, 9, Int($iUM))

		Case $__LOCONST_CONVERT_PT_UM
			; 1 pt = 35 uM
			Return ($nValue = 0) ? (SetError($__LO_STATUS_SUCCESS, 10, 0)) : (SetError($__LO_STATUS_SUCCESS, 10, Round(($nValue * 35.2778))))

		Case $__LOCONST_CONVERT_UM_PT
			Return ($nValue = 0) ? (SetError($__LO_STATUS_SUCCESS, 11, 0)) : (SetError($__LO_STATUS_SUCCESS, 11, Round(($nValue / 35.2778), 2)))

		Case Else
			Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndSwitch
EndFunc   ;==>__LOCalc_UnitConvert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_VarsAreDefault
; Description ...: Tests whether all input parameters are equal to Default keyword.
; Syntax ........: __LOCalc_VarsAreDefault($vVar1[, $vVar2 = Default[, $vVar3 = Default[, $vVar4 = Default[, $vVar5 = Default[, $vVar6 = Default[, $vVar7 = Default[, $vVar8 = Default]]]]]]])
; Parameters ....: $vVar1               - a variant value.
;                  $vVar2               - [optional] a variant value. Default is Default.
;                  $vVar3               - [optional] a variant value. Default is Default.
;                  $vVar4               - [optional] a variant value. Default is Default.
;                  $vVar5               - [optional] a variant value. Default is Default.
;                  $vVar6               - [optional] a variant value. Default is Default.
;                  $vVar7               - [optional] a variant value. Default is Default.
;                  $vVar8               - [optional] a variant value. Default is Default.
; Return values .: Success: Boolean
;				   Failure: False
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If All parameters are Equal to Default, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_VarsAreDefault($vVar1, $vVar2 = Default, $vVar3 = Default, $vVar4 = Default, $vVar5 = Default, $vVar6 = Default, $vVar7 = Default, $vVar8 = Default)
	Local $bAllDefault1, $bAllDefault2
	$bAllDefault1 = (($vVar1 = Default) And ($vVar2 = Default) And ($vVar3 = Default) And ($vVar4 = Default)) ? (True) : (False)
	$bAllDefault2 = (($vVar5 = Default) And ($vVar6 = Default) And ($vVar7 = Default) And ($vVar8 = Default)) ? (True) : (False)
	Return ($bAllDefault1 And $bAllDefault2) ? (True) : (False)
EndFunc   ;==>__LOCalc_VarsAreDefault

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_VarsAreNull
; Description ...: Tests whether all input parameters are equal to Null keyword.
; Syntax ........: __LOCalc_VarsAreNull($vVar1[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null[, $vVar9 = Null[, $vVar10 = Null[, $vVar11 = Null[, $vVar12 = Null]]]]]]]]]]])
; Parameters ....: $vVar1               - a variant value.
;                  $vVar2               - [optional] a variant value. Default is Null.
;                  $vVar3               - [optional] a variant value. Default is Null.
;                  $vVar4               - [optional] a variant value. Default is Null.
;                  $vVar5               - [optional] a variant value. Default is Null.
;                  $vVar6               - [optional] a variant value. Default is Null.
;                  $vVar7               - [optional] a variant value. Default is Null.
;                  $vVar8               - [optional] a variant value. Default is Null.
;                  $vVar9               - [optional] a variant value. Default is Null.
;                  $vVar10              - [optional] a variant value. Default is Null.
;                  $vVar11              - [optional] a variant value. Default is Null.
;                  $vVar12              - [optional] a variant value. Default is Null.
; Return values .: Success: Boolean
;				   Failure: False
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If All parameters are Equal to Null, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_VarsAreNull($vVar1, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null, $vVar9 = Null, $vVar10 = Null, $vVar11 = Null, $vVar12 = Null)
	Local $bAllNull1, $bAllNull2, $bAllNull3
	$bAllNull1 = (($vVar1 = Null) And ($vVar2 = Null) And ($vVar3 = Null) And ($vVar4 = Null)) ? (True) : (False)
	If (@NumParams <= 4) Then Return ($bAllNull1) ? (True) : (False)
	$bAllNull2 = (($vVar5 = Null) And ($vVar6 = Null) And ($vVar7 = Null) And ($vVar8 = Null)) ? (True) : (False)
	If (@NumParams <= 8) Then Return ($bAllNull1 And $bAllNull2) ? (True) : (False)
	$bAllNull3 = (($vVar9 = Null) And ($vVar10 = Null) And ($vVar11 = Null) And ($vVar12 = Null)) ? (True) : (False)
	Return ($bAllNull1 And $bAllNull2 And $bAllNull3) ? (True) : (False)
EndFunc   ;==>__LOCalc_VarsAreNull

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_VersionCheck
; Description ...: Test if the currently installed LibreOffice version is high enough to support a certain function.
; Syntax ........: __LOCalc_VersionCheck($fRequiredVersion)
; Parameters ....: $fRequiredVersion            - a floating point value. The version of LibreOffice required.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $fRequiredVersion not a Number.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Current L.O. Version.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. If the Current L.O. version is higher than or equal to the required version, then True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_VersionCheck($fRequiredVersion)
	Local Static $sCurrentVersion = _LOCalc_VersionGet(True, False)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, False)
	Local Static $fCurrentVersion = Number($sCurrentVersion)

	If Not IsNumber($fRequiredVersion) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, False)

	Return SetError($__LO_STATUS_SUCCESS, 1, ($fCurrentVersion >= $fRequiredVersion) ? (True) : (False))
EndFunc   ;==>__LOCalc_VersionCheck
