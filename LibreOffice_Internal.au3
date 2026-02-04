#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1

#include-once

#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Internal functions for interacting with Libre Office.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __LO_AddTo1DArray
; __LO_ArrayFill
; __LO_CreateStruct
; __LO_DeleteTempReg
; __LO_InternalComErrorHandler
; __LO_IntIsBetween
; __LO_NumIsBetween
; __LO_ServiceManager
; __LO_SetPortableServiceManager
; __LO_SetPropertyValue
; __LO_StylesGetNames
; __LO_VarsAreNull
; __LO_VersionCheck
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_AddTo1DArray
; Description ...: Add data to a 1 Dimensional array.
; Syntax ........: __LO_AddTo1DArray(ByRef $aArray, $vData[, $bCountInFirst = False])
; Parameters ....: $aArray              - [in/out] an array of unknowns. The Array to directly add data to. Array will be directly modified.
;                  $vData               - a variant value. The Data to add to the Array.
;                  $bCountInFirst       - [optional] a boolean value. Default is False. If True the first element of the array is a count of contained elements.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $aArray not an Array
;                  @Error 1 @Extended 2 Return 0 = $bCountinFirst not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $aArray contains too many columns.
;                  @Error 1 @Extended 4 Return 0 = $aArray[0] contains non-Integer data or is not empty, and $bCountInFirst is called with True.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Array item was successfully added.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_AddTo1DArray(ByRef $aArray, $vData, $bCountInFirst = False)
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
EndFunc   ;==>__LO_AddTo1DArray

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_ArrayFill
; Description ...: Fill an Array with data.
; Syntax ........: __LO_ArrayFill(ByRef $aArrayToFill[, $vVar1 = Null[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null[, $vVar9 = Null[, $vVar10 = Null[, $vVar11 = Null[, $vVar12 = Null[, $vVar13 = Null[, $vVar14 = Null[, $vVar15 = Null[, $vVar16 = Null[, $vVar17 = Null[, $vVar18 = Null[, $vVar19 = Null[, $vVar20 = Null[, $vVar21 = Null[, $vVar22 = Null[, $vVar23 = Null[, $vVar24 = Null[, $vVar25 = Null[, $vVar26 = Null[, $vVar27 = Null[, $vVar28 = Null[, $vVar29 = Null[, $vVar30 = Null[, $vVar31 = Null[, $vVar32 = Null]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]])
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
;                  $vVar19              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar20              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar21              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar22              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar23              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar24              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar25              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar26              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar27              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar28              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar29              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar30              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar31              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar32              - [optional] a variant value. Default is Null. The Data to add to the Array.
; Return values .: None
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call only how many you parameters you need to add to the Array. Automatically resizes the Array if it is the incorrect size.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_ArrayFill(ByRef $aArrayToFill, $vVar1 = Null, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null, $vVar9 = Null, $vVar10 = Null, $vVar11 = Null, $vVar12 = Null, $vVar13 = Null, $vVar14 = Null, $vVar15 = Null, $vVar16 = Null, $vVar17 = Null, $vVar18 = Null, $vVar19 = Null, $vVar20 = Null, $vVar21 = Null, $vVar22 = Null, $vVar23 = Null, $vVar24 = Null, $vVar25 = Null, $vVar26 = Null, $vVar27 = Null, $vVar28 = Null, $vVar29 = Null, $vVar30 = Null, $vVar31 = Null, $vVar32 = Null)
	#forceref $vVar1, $vVar2, $vVar3, $vVar4, $vVar5, $vVar6, $vVar7, $vVar8, $vVar9, $vVar10, $vVar11, $vVar12, $vVar13, $vVar14, $vVar15, $vVar16, $vVar17, $vVar18, $vVar19, $vVar20, $vVar21, $vVar22, $vVar23, $vVar24, $vVar25, $vVar26, $vVar27, $vVar28, $vVar29, $vVar30, $vVar31, $vVar32

	If UBound($aArrayToFill) < (@NumParams - 1) Then ReDim $aArrayToFill[@NumParams - 1]
	For $i = 0 To @NumParams - 2
		$aArrayToFill[$i] = Eval("vVar" & $i + 1)
	Next
EndFunc   ;==>__LO_ArrayFill

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_CreateStruct
; Description ...: Retrieves a Struct.
; Syntax ........: __LO_CreateStruct($sStructName)
; Parameters ....: $sStructName         - a string value. Name of structure to create.
; Return values .: Success: Structure.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sStructName not a string
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object
;                  @Error 2 @Extended 2 Return 0 = Error creating requested structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Structure = Success. Property Structure Returned
; Author ........: mLipok
; Modified ......: donnyh13 - Added error checking.
; Remarks .......: From WriterDemo.au3 as modified by mLipok from WriterDemo.vbs found in the LibreOffice SDK examples.
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/204665-libreopenoffice-writer/?do=findComment&comment=1471711
; Example .......: No
; ===============================================================================================================================
Func __LO_CreateStruct($sStructName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $tStruct

	If Not IsString($sStructName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tStruct = $oServiceManager.Bridge_GetStruct($sStructName)
	If Not IsObj($tStruct) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $tStruct)
EndFunc   ;==>__LO_CreateStruct

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_DeleteTempReg
; Description ...: Delete Temporary Registry entries used for connecting to Portable LO.
; Syntax ........: __LO_DeleteTempReg([$asRegKeys = Null])
; Parameters ....: $asRegKeys           - [optional] an array of strings. Default is Null. An array of Registry keys to Delete.
; Return values .: Success: 1, 2
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $asRegKeys not an Array.
;                  --Processing Errors--
;                  @Error 3 @Extended ? Return 0 = Error Deleting Registry key. @Extended set to number of errors encountered.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully stored Registry keys to delete.
;                  @Error 0 @Extended 0 Return 2 = Success. Successfully deleted Registry keys.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_DeleteTempReg($asRegKeys = Null)
	Local Static $asStaticKeys[0]
	Local $iError = 0
	Local Const $sHKCU = (@OSArch = "X86") ? ("HKCU") : ("HKCU64")

	If ($asRegKeys <> Null) Then
		If Not IsArray($asRegKeys) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		$asStaticKeys = $asRegKeys

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndIf

	For $sKey In $asStaticKeys
		RegDelete($sHKCU & $sKey)
		$iError = (@error > 0) ? ($iError + 1) : ($iError)
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROCESSING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 2))
EndFunc   ;==>__LO_DeleteTempReg

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LO_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _LO_ComError_UserFunction(Default)
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
EndFunc   ;==>__LO_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_IntIsBetween
; Description ...: Test whether an input is an Integer and is between two Integers.
; Syntax ........: __LO_IntIsBetween($iTest, $iMin, $iMax[, $vNot = ""[, $vIncl = ""]])
; Parameters ....: $iTest               - an integer value. The Value to test.
;                  $iMin                - an integer value. The minimum $iTest can be.
;                  $iMax                - [optional] an integer value. Default is 0. The maximum $iTest can be.
;                  $vNot                - [optional] a variant value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers inside the min/max range that are not allowed.
;                  $vIncl               - [optional] a variant value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers Outside the min/max range that are allowed.
; Return values .: Success: Boolean
;                  Failure: False and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return Boolean = $iTest not an Integer.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the input is between Min and Max or is an allowed number, and not one of the disallowed numbers, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_IntIsBetween($iTest, $iMin, $iMax = 0, $vNot = "", $vIncl = "")
	If Not IsInt($iTest) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, False)

	Switch @NumParams
		Case 2

			Return SetError($__LO_STATUS_SUCCESS, 0, ($iTest < $iMin) ? (False) : (True))

		Case 3

			Return SetError($__LO_STATUS_SUCCESS, 0, (($iTest < $iMin) Or ($iTest > $iMax)) ? (False) : (True))

		Case 4, 5
			If IsString($vNot) Then
				If StringInStr(":" & $vNot & ":", ":" & $iTest & ":") Then Return SetError($__LO_STATUS_SUCCESS, 0, False)

			ElseIf IsInt($vNot) Then
				If ($iTest = $vNot) Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
			EndIf

			If (($iTest >= $iMin) And ($iTest <= $iMax)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

			If @NumParams = 5 Then ContinueCase

			Return SetError($__LO_STATUS_SUCCESS, 0, False)

		Case Else
			If IsString($vIncl) Then
				If StringInStr(":" & $vIncl & ":", ":" & $iTest & ":") Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

			ElseIf IsInt($vIncl) Then
				If ($iTest = $vIncl) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
			EndIf

			Return SetError($__LO_STATUS_SUCCESS, 0, False)
	EndSwitch
EndFunc   ;==>__LO_IntIsBetween

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_NumIsBetween
; Description ...: Test whether an input is a Number and is between two Numbers.
; Syntax ........: __LO_NumIsBetween($nTest, $nMin, $nMax[, $snNot = ""[, $snIncl = Default]])
; Parameters ....: $nTest               - a general number value. The Value to test.
;                  $nMin                - a general number value. The minimum $iTest can be.
;                  $nMax                - a general number value. The maximum $iTest can be.
;                  $snNot               - [optional] a string value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers inside the min/max range that are not allowed.
;                  $snIncl              - [optional] a string value. Default is Default. Can be a single number, or a String of numbers separated by ":". Defines numbers Outside the min/max range that are allowed.
; Return values .: Success: Boolean
;                  Failure: False
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the input is between Min and Max or is an allowed number, and not one of the disallowed numbers, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_NumIsBetween($nTest, $nMin, $nMax, $snNot = "", $snIncl = Default)
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
EndFunc   ;==>__LO_NumIsBetween

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_ServiceManager
; Description ...: Set or Retrieve a stored Service Manager Object for use in the UDF.
; Syntax ........: __LO_ServiceManager([$oServiceManager = Null[, $bPortable = Null]])
; Parameters ....: $oServiceManager     - [optional] an object. Default is Null. A ServiceManager Object. Typically this is used to store a Portable Service Manager Object.
;                  $bPortable           - [optional] a boolean value. Default is Null. If True, a Portable LibreOffice ServiceManager will be stored.
; Return values .: Success: 1 or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bPortable not a Boolean.
;                  @Error 1 @Extended 2 Return 0 = $oServiceManager not an Object.
;                  @Error 1 @Extended 3 Return 0 =Object called in $oServiceManager not a ServiceManager Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a ServiceManager.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully cleared stored ServiceManager.
;                  @Error 0 @Extended 0 Return Object = Success. Returning ServiceManager Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_ServiceManager($oServiceManager = Null, $bPortable = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Static $oStaticServiceManager
	Local Static $bIsPortable = False

	If ($bPortable <> Null) Then
		If Not IsBool($bPortable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		$bIsPortable = $bPortable
	EndIf

	If ($oServiceManager <> Null) Then
		If ($oServiceManager = Default) Then ; Clear the saved Service Manager. This could be used in the case of switching from portable to installed.
			$oStaticServiceManager = Null

		Else
			If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
			If Not $oServiceManager.supportsService("com.sun.star.lang.ServiceManager") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$oStaticServiceManager = $oServiceManager
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndIf

	If IsObj($oStaticServiceManager) Then ; Test if the Object is still valid.
		If Not IsBool($oStaticServiceManager.supportsService("com.sun.star.lang.ServiceManager")) Then $oStaticServiceManager = Null
	EndIf

	If Not IsObj($oStaticServiceManager) Then
		If $bIsPortable Then
			; Try to create the ServiceManager again for the portable version.
			__LO_SetPortableServiceManager()

		Else ; Create a ServiceManager, for the installed version.
			$oStaticServiceManager = ObjCreate("com.sun.star.ServiceManager")
		EndIf
	EndIf

	If Not IsObj($oStaticServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 1, $oStaticServiceManager)
EndFunc   ;==>__LO_ServiceManager

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_SetPortableServiceManager
; Description ...: Create and store a Portable LibreOffice ServiceManager Object.
; Syntax ........: __LO_SetPortableServiceManager([$sPortableLO_Path = Null])
; Parameters ....: $sPortableLO_Path    - [optional] a string value. Default is Null. A path to the Portable LibreOffice soffice.exe file.
; Return values .: Success: 1, 2
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sPortableLO_Path not a String.
;                  @Error 1 @Extended 2 Return 0 = Path called in $sPortableLO_Path doesn't exist.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.bridge.UnoUrlResolver" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = No stored Portable LibreOffice path.
;                  @Error 3 @Extended 2 Return 0 = Stored Portable LibreOffice path no longer exists.
;                  @Error 3 @Extended 3 Return 0 = Failed to add temporary Registry keys.
;                  @Error 3 @Extended 4 Return 0 = Failed to start Portable LibreOffice in listening mode.
;                  @Error 3 @Extended 5 Return 0 = Portable LibreOffice failed to start in listening mode
;                  @Error 3 @Extended 6 Return 0 = Failed to connect to Portable LibreOffice.
;                  @Error 3 @Extended 7 Return 0 = Failed to retrieve ServiceManager from Portable LibreOffice.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Created and stored Portable LibreOffice ServiceManager.
;                  @Error 0 @Extended 0 Return 2 = Success. Cleared stored Portable LibreOffice path.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the COM Error "Binary URP bridge already disposed" is encountered, any running soffice.exe/soffice.bin processes need to be closed via TaskManager.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_SetPortableServiceManager($sPortableLO_Path = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Static $sStaticPortablePath = ""
	Local Static $bTempReg = False
	Local $iError = 0, $iPort = 2002, $iSocket, $iErrorRet
	Local $hTimer
	Local $oReg_ServiceManager, $oURLResolver, $oCompContext, $oPortable_ServiceManager
	Local Const $__eReg_LocalServer32 = 2 ; Make sure to sync this.
	Local Enum $__eReg_KeyName, $__eReg_ValueName, $__eReg_Type, $__eReg_Value
	Local Const $sHKCU = (@OSArch = "X86") ? ("HKCU") : ("HKCU64")
	Local Const $sHKLM = (@OSArch = "X86") ? ("HKLM") : ("HLM64")
	Local $asRegKeysMain[2] = ["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}", _
			"\Software\Classes\com.sun.star.ServiceManager"]
	Local $asRegKeys[11][4] = [["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}", "", "REG_SZ", "LibreOffice Service Manager (Ver 1.0)"], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}", "AppID", "REG_SZ", "{82154420-0FBF-11d4-8313-005004526AB4}"], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}\LocalServer32", "", "REG_SZ", ""], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}\NotInsertable", "", "REG_SZ", ""], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}\ProgID", "", "REG_SZ", "com.sun.star.ServiceManager.1"], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}\Programmable", "", "REG_SZ", ""], _
			["\Software\Classes\CLSID\{82154420-0FBF-11d4-8313-005004526AB4}\VersionIndependentProgID", "", "REG_SZ", "com.sun.star.ServiceManager"], _
			["\Software\Classes\com.sun.star.ServiceManager", "", "REG_SZ", "LibreOffice Service Manager"], _
			["\Software\Classes\com.sun.star.ServiceManager\CLSID", "", "REG_SZ", "{82154420-0FBF-11d4-8313-005004526AB4}"], _
			["\Software\Classes\com.sun.star.ServiceManager\CurVer", "", "REG_SZ", "com.sun.star.ServiceManager.1"], _
			["\Software\Classes\com.sun.star.ServiceManager\NotInsertable", "", "REG_SZ", ""]]

	If ($sPortableLO_Path <> Null) Then
		If Not IsString($sPortableLO_Path) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		If ($sPortableLO_Path <> "") Then
			If Not FileExists($sPortableLO_Path) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

			$sStaticPortablePath = $sPortableLO_Path

		Else
			$sStaticPortablePath = $sPortableLO_Path
			__LO_ServiceManager(Default, False) ; Clear any stored ServiceManager, and set Boolean for Portable to False.

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf
	EndIf

	If ($sStaticPortablePath = "") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not FileExists($sStaticPortablePath) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Check to make sure the stored path to the LO File is still good.

	$asRegKeys[$__eReg_LocalServer32][$__eReg_Value] = $sStaticPortablePath & " --nodefault --nologo"

	RegRead($sHKLM & $asRegKeysMain[1], "") ; Classes are shared, so I only need to check one place.
	If @error Then ; Seems like no LibreOffice is installed, add temp Registry.
		For $i = 0 To UBound($asRegKeys) - 1
			RegWrite($sHKCU & $asRegKeys[$i][$__eReg_KeyName], $asRegKeys[$i][$__eReg_ValueName], $asRegKeys[$i][$__eReg_Type], $asRegKeys[$i][$__eReg_Value])
			$iError = (@error > 0) ? ($iError + 1) : ($iError)
		Next

		__LO_DeleteTempReg($asRegKeysMain) ; Set array of main Temp keys to delete.
		If ($iError > 0) Then ; If there was an error writing the Reg Keys, delete any that were written and return.
			__LO_DeleteTempReg()

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		EndIf

		OnAutoItExitRegister("__LO_DeleteTempReg")
		$bTempReg = True
	EndIf

	$oReg_ServiceManager = ObjCreate("com.sun.star.ServiceManager") ; Create a ServiceManager from the one registered in Registry.
	If Not IsObj($oReg_ServiceManager) Then
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf

		Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	EndIf

	TCPStartup()

	; Find an unused port.
	While (TCPConnect("127.0.0.1", $iPort) > 0)
		Sleep(10)
		$iPort += 1
	WEnd

	; Start LibreOffice Portable in listening mode
	Run('"' & $sStaticPortablePath & '" --headless --norestore --nologo --accept="socket,host=127.0.0.1,port=' & $iPort & ',tcpNoDelay=1;urp;"', "", @SW_HIDE)
	If (@error <> 0) Then
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf
		TCPShutdown()

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$hTimer = TimerInit()

	; Wait until LO is initialized and listening.
	Do
		Sleep(10)
		$iSocket = TCPConnect("127.0.0.1", $iPort)
		$iErrorRet = @error
	Until (($iErrorRet = 0) And ($iSocket > 0)) Or (TimerDiff($hTimer) > 15000) ; Timeout in 15 seconds.
	TCPShutdown()

	If ($iErrorRet > 0) Then ; Error initializing LO.
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
	EndIf

	; Create the URL Resolver.
	$oURLResolver = $oReg_ServiceManager.createInstance("com.sun.star.bridge.UnoUrlResolver")
	If Not IsObj($oURLResolver) Then
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf

		Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	EndIf

	$oCompContext = $oURLResolver.resolve("uno:socket,host=localhost,port=" & $iPort & ",tcpNoDelay=1;urp;StarOffice.ComponentContext")

	If Not IsObj($oCompContext) Then
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
	EndIf

	$oPortable_ServiceManager = $oCompContext.getServiceManager() ; Get ServiceManager of Portable LO.
	If Not IsObj($oPortable_ServiceManager) Then
		If $bTempReg Then
			__LO_DeleteTempReg()
			If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
			$bTempReg = False
		EndIf

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)
	EndIf

	__LO_ServiceManager($oPortable_ServiceManager, True) ; Set the stored ServiceManager.

	If $bTempReg Then ; Clean up Temp Registry.
		__LO_DeleteTempReg()
		If (@error = 0) Then OnAutoItExitUnRegister("__LO_DeleteTempReg")
		$bTempReg = False
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LO_SetPortableServiceManager

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_SetPropertyValue
; Description ...: Creates a property value struct object.
; Syntax ........: __LO_SetPropertyValue($sName, $vValue)
; Parameters ....: $sName               - a string value. Property name.
;                  $vValue              - a variant value. Property value.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sName not a string
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Properties Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Property Object Returned
; Author ........: Leagnus, GMK
; Modified ......: donnyh13 - added CreateStruct function. Modified variable names.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_SetPropertyValue($sName, $vValue)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tProperties

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tProperties = __LO_CreateStruct("com.sun.star.beans.PropertyValue")
	If @error Or Not IsObj($tProperties) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProperties.Name = $sName
	$tProperties.Value = $vValue

	Return SetError($__LO_STATUS_SUCCESS, 0, $tProperties)
EndFunc   ;==>__LO_SetPropertyValue

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_StylesGetNames
; Description ...: Retrieve an Array of Style names available.
; Syntax ........: __LO_StylesGetNames(ByRef $oDoc, $sStyleFamily[, $bUserOnly = False[, $bAppliedOnly = False[, $bDisplayName = False]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LO*_DocOpen, _LO*_DocConnect, or _LO*_DocCreate function.
;                  $sStyleFamily        - a string value. The Style type to retrieve names for.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True, only user-created Styles are returned.
;                  $bAppliedOnly        - [optional] a boolean value. Default is False. If True, only applied styles are returned.
;                  $bDisplayName        - [optional] a boolean value. Default is False. If True, the style name displayed in the UI (Display Name), instead of the programmatic style name, is returned. See remarks.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sStyleFamily not a String.
;                  @Error 1 @Extended 3 Return 0 = $bUserOnly not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bAppliedOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bDisplayName not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = Style family called in $sStyleFamily doesn't exist.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve called Style family Object.
;                  --Success--
;                  @Error 0 @Extended 1 Return Array = Success. An Array containing all Styles matching the called parameters. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If Only a Document object is called, all available Paragraph styles will be returned.
;                  If Both $bUserOnly and $bAppliedOnly are called with True, only User-Created styles that are applied are returned.
;                  Calling $bDisplayName with True will return a list of Style names, as the user sees them in the UI, in the same order as they are returned if $bDisplayName is False. It is best not to use these when setting Paragraph Styling.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_StylesGetNames(ByRef $oDoc, $sStyleFamily, $bUserOnly = False, $bAppliedOnly = False, $bDisplayName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oStyles
	Local $asStyles[0]
	Local $iCount = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sStyleFamily) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bDisplayName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not $oDoc.StyleFamilies.hasByName($sStyleFamily) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oStyles = $oDoc.StyleFamilies.getByName($sStyleFamily)
	If Not IsObj($oStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $asStyles[$oStyles.getCount()]

	If $bUserOnly And $bAppliedOnly Then
		For $i = 0 To $oStyles.getCount() - 1
			If ($oStyles.getByIndex($i).isUserDefined()) And ($oStyles.getByIndex($i).isInUse()) Then
				If $bDisplayName Then
					$asStyles[$iCount] = $oStyles.getByIndex($i).DisplayName()

				Else
					$asStyles[$iCount] = $oStyles.getByIndex($i).Name()
				EndIf
				$iCount += 1
			EndIf

			Sleep((IsInt($i / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		ReDim $asStyles[$iCount]

	ElseIf $bUserOnly Then
		For $i = 0 To $oStyles.getCount() - 1
			If $oStyles.getByIndex($i).isUserDefined() Then
				If $bDisplayName Then
					$asStyles[$iCount] = $oStyles.getByIndex($i).DisplayName()

				Else
					$asStyles[$iCount] = $oStyles.getByIndex($i).Name()
				EndIf
				$iCount += 1
			EndIf

			Sleep((IsInt($i / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		ReDim $asStyles[$iCount]

	ElseIf $bAppliedOnly Then
		For $i = 0 To $oStyles.getCount() - 1
			If $oStyles.getByIndex($i).isInUse() Then
				If $bDisplayName Then
					$asStyles[$iCount] = $oStyles.getByIndex($i).DisplayName()

				Else
					$asStyles[$iCount] = $oStyles.getByIndex($i).Name()
				EndIf
				$iCount += 1
			EndIf

			Sleep((IsInt($i / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		ReDim $asStyles[$iCount]

	Else ; Get all Styles.
		For $i = 0 To $oStyles.getCount() - 1
			If $bDisplayName Then
				$asStyles[$i] = $oStyles.getByIndex($i).DisplayName()

			Else
				$asStyles[$i] = $oStyles.getByIndex($i).Name()
			EndIf

			Sleep((IsInt($i / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 1, $asStyles)
EndFunc   ;==>__LO_StylesGetNames

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_VarsAreNull
; Description ...: Tests whether all input parameters are equal to Null keyword.
; Syntax ........: __LO_VarsAreNull($vVar1[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null[, $vVar9 = Null[, $vVar10 = Null[, $vVar11 = Null[, $vVar12 = Null[, $vVar13 = Null[, $vVar14 = Null[, $vVar15 = Null[, $vVar16 = Null[, $vVar17 = Null[, $vVar18 = Null[, $vVar19 = Null[, $vVar20 = Null[, $vVar21 = Null[, $vVar22 = Null[, $vVar23 = Null[, $vVar24 = Null[, $vVar25 = Null[, $vVar26 = Null[, $vVar27 = Null[, $vVar28 = Null[, $vVar29 = Null[, $vVar30 = Null[, $vVar31 = Null[, $vVar32 = Null]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]])
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
;                  $vVar13              - [optional] a variant value. Default is Null.
;                  $vVar14              - [optional] a variant value. Default is Null.
;                  $vVar15              - [optional] a variant value. Default is Null.
;                  $vVar16              - [optional] a variant value. Default is Null.
;                  $vVar17              - [optional] a variant value. Default is Null.
;                  $vVar18              - [optional] a variant value. Default is Null.
;                  $vVar19              - [optional] a variant value. Default is Null.
;                  $vVar20              - [optional] a variant value. Default is Null.
;                  $vVar21              - [optional] a variant value. Default is Null.
;                  $vVar22              - [optional] a variant value. Default is Null.
;                  $vVar23              - [optional] a variant value. Default is Null.
;                  $vVar24              - [optional] a variant value. Default is Null.
;                  $vVar25              - [optional] a variant value. Default is Null.
;                  $vVar26              - [optional] a variant value. Default is Null.
;                  $vVar27              - [optional] a variant value. Default is Null.
;                  $vVar28              - [optional] a variant value. Default is Null.
;                  $vVar29              - [optional] a variant value. Default is Null.
;                  $vVar30              - [optional] a variant value. Default is Null.
;                  $vVar31              - [optional] a variant value. Default is Null.
;                  $vVar32              - [optional] a variant value. Default is Null.
; Return values .: Success: Boolean
;                  Failure: False
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If All parameters are Equal to Null, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_VarsAreNull($vVar1, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null, $vVar9 = Null, $vVar10 = Null, $vVar11 = Null, $vVar12 = Null, $vVar13 = Null, $vVar14 = Null, $vVar15 = Null, $vVar16 = Null, $vVar17 = Null, $vVar18 = Null, $vVar19 = Null, $vVar20 = Null, $vVar21 = Null, $vVar22 = Null, $vVar23 = Null, $vVar24 = Null, $vVar25 = Null, $vVar26 = Null, $vVar27 = Null, $vVar28 = Null, $vVar29 = Null, $vVar30 = Null, $vVar31 = Null, $vVar32 = Null)
	Local $bAllNull1, $bAllNull2, $bAllNull3, $bAllNull4, $bAllNull5, $bAllNull6, $bAllNull7, $bAllNull8
	$bAllNull1 = (($vVar1 = Null) And ($vVar2 = Null) And ($vVar3 = Null) And ($vVar4 = Null)) ? (True) : (False)
	If (@NumParams <= 4) Then Return SetError($__LO_STATUS_SUCCESS, 0, ($bAllNull1) ? (True) : (False))

	$bAllNull2 = (($vVar5 = Null) And ($vVar6 = Null) And ($vVar7 = Null) And ($vVar8 = Null)) ? (True) : (False)
	If (@NumParams <= 8) Then Return SetError($__LO_STATUS_SUCCESS, 0, (($bAllNull1) And ($bAllNull2)) ? (True) : (False))

	$bAllNull3 = (($vVar9 = Null) And ($vVar10 = Null) And ($vVar11 = Null) And ($vVar12 = Null)) ? (True) : (False)
	If (@NumParams <= 12) Then Return SetError($__LO_STATUS_SUCCESS, 0, (($bAllNull1) And ($bAllNull2) And ($bAllNull3)) ? (True) : (False))

	$bAllNull4 = (($vVar13 = Null) And ($vVar14 = Null) And ($vVar15 = Null) And ($vVar16 = Null)) ? (True) : (False)
	If (@NumParams <= 16) Then Return SetError($__LO_STATUS_SUCCESS, 0, (($bAllNull1) And ($bAllNull2) And ($bAllNull3) And ($bAllNull4)) ? (True) : (False))

	$bAllNull5 = (($vVar17 = Null) And ($vVar18 = Null) And ($vVar19 = Null) And ($vVar20 = Null)) ? (True) : (False)
	If (@NumParams <= 20) Then Return SetError($__LO_STATUS_SUCCESS, 0, (($bAllNull1) And ($bAllNull2) And ($bAllNull3) And ($bAllNull4) And ($bAllNull5)) ? (True) : (False))

	$bAllNull6 = (($vVar21 = Null) And ($vVar22 = Null) And ($vVar23 = Null) And ($vVar24 = Null)) ? (True) : (False)
	If (@NumParams <= 24) Then Return SetError($__LO_STATUS_SUCCESS, 0, (($bAllNull1) And ($bAllNull2) And ($bAllNull3) And ($bAllNull4) And ($bAllNull5) And ($bAllNull6)) ? (True) : (False))

	$bAllNull7 = (($vVar25 = Null) And ($vVar26 = Null) And ($vVar27 = Null) And ($vVar28 = Null)) ? (True) : (False)
	If (@NumParams <= 28) Then Return SetError($__LO_STATUS_SUCCESS, 0, (($bAllNull1) And ($bAllNull2) And ($bAllNull3) And ($bAllNull4) And ($bAllNull5) And ($bAllNull6) And ($bAllNull7)) ? (True) : (False))

	$bAllNull8 = (($vVar29 = Null) And ($vVar30 = Null) And ($vVar31 = Null) And ($vVar32 = Null)) ? (True) : (False)

	Return SetError($__LO_STATUS_SUCCESS, 0, (($bAllNull1) And ($bAllNull2) And ($bAllNull3) And ($bAllNull4) And ($bAllNull5) And ($bAllNull6) And ($bAllNull7) And ($bAllNull8)) ? (True) : (False))
EndFunc   ;==>__LO_VarsAreNull

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LO_VersionCheck
; Description ...: Test if the currently installed LibreOffice version is high enough to support a certain function.
; Syntax ........: __LO_VersionCheck($fRequiredVersion)
; Parameters ....: $fRequiredVersion    - a floating point value. The version of LibreOffice required.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $fRequiredVersion not a Number.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Current L.O. Version.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the Current L.O. version is greater than or equal to the required version, then True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LO_VersionCheck($fRequiredVersion)
	Local Static $sCurrentVersion = _LO_VersionGet(True, False)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, False)

	Local Static $fCurrentVersion = Number($sCurrentVersion)

	If Not IsNumber($fRequiredVersion) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, False)

	Return SetError($__LO_STATUS_SUCCESS, 1, ($fCurrentVersion >= $fRequiredVersion) ? (True) : (False))
EndFunc   ;==>__LO_VersionCheck
