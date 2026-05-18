#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Constants.au3"
#include "LibreOfficeCalc_Internal.au3"

; Other includes for Calc

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Functions used for creating, modifying and retrieving data for use in various functions in LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_ComError_UserFunction
; _LOCalc_FilterDescriptorCreate
; _LOCalc_FilterDescriptorModify
; _LOCalc_FilterFieldCreate
; _LOCalc_FilterFieldModify
; _LOCalc_FontExists
; _LOCalc_FontsGetNames
; _LOCalc_FormatKeyCreate
; _LOCalc_FormatKeyDelete
; _LOCalc_FormatKeyExists
; _LOCalc_FormatKeyGetStandard
; _LOCalc_FormatKeyGetString
; _LOCalc_FormatKeysGetList
; _LOCalc_SearchDescriptorCreate
; _LOCalc_SearchDescriptorModify
; _LOCalc_SearchDescriptorSimilarityModify
; _LOCalc_SortFieldCreate
; _LOCalc_SortFieldModify
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LOCalc_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
; Parameters ....: $vUserFunction       - [optional] a Function or Keyword. Default is Default. Accepts a Function, or the Keyword Default and Null. If called with a User function, the function may have up to 5 required parameters.
;                  $vParam1             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam2             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam3             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam4             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam5             - [optional] Default is Null. Any optional parameter to be called with the user function.
; Return values .: Success: 1 or UserFunction.
;                  @Error: 0, @Extended: 0, Return: 1 = Successfully set the UserFunction.
;                  @Error: 0, @Extended: 0, Return: 2 = Successfully cleared the set UserFunction.
;                  @Error: 0, @Extended: 0, Return: Function = Returning the set UserFunction.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $vUserFunction Not a Function, or Default keyword, or Null Keyword.
; Author ........: mLipok
; Modified ......: donnyh13 - Added a clear UserFunction without error option. Also added parameters option.
; Remarks .......: The first parameter passed to the User function will always be the COM Error object. See below.
;                  Every COM Error will be passed to that function. The user can then read the following properties. (As Found in the COM Reference section in Autoit HelpFile.) Using the first parameter in the UserFunction.
;                  For Example MyFunc($oMyError)
;                  - $oMyError.number The Windows HRESULT value from a COM call
;                  - $oMyError.windescription The FormatWinError() text derived from .number
;                  - $oMyError.source Name of the Object generating the error (contents from ExcepInfo.source)
;                  - $oMyError.description Source Object's description of the error (contents from ExcepInfo.description)
;                  - $oMyError.helpfile Source Object's help file for the error (contents from ExcepInfo.helpfile)
;                  - $oMyError.helpcontext Source Object's help file context id number (contents from ExcepInfo.helpcontext)
;                  - $oMyError.lastdllerror The number returned from GetLastError()
;                  - $oMyError.scriptline The script line on which the error was generated
;                  - NOTE: Not all properties will necessarily contain data, some will be blank.
;                  If MsgBox or ConsoleWrite functions are passed to this function, the error details will be displayed using that function automatically.
;                  If called with Default keyword, the current UserFunction, if set, will be returned.
;                  If called with Null keyword, the currently set UserFunction is cleared and only the internal ComErrorHandler will be called for COM Errors.
;                  The stored UserFunction (besides MsgBox and ConsoleWrite) will be called as follows: UserFunc($oComError,$vParam1,$vParam2,$vParam3,$vParam4,$vParam5)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
	#forceref $vParam1, $vParam2, $vParam3, $vParam4, $vParam5

	; If user does not set a function, UDF must use internal function to avoid AutoItError.
	Local Static $vUserFunction_Static = Default
	Local $avUserFuncWParams[@NumParams]

	If $vUserFunction = Default Then
		; just return stored static User Function variable

		Return SetError($__LO_STATUS_SUCCESS, 0, $vUserFunction_Static)

	ElseIf IsFunc($vUserFunction) Then
		; If User called Parameters, then add to array.
		If @NumParams > 1 Then
			$avUserFuncWParams[0] = $vUserFunction
			For $i = 1 To @NumParams - 1
				$avUserFuncWParams[$i] = Eval("vParam" & $i)
				; set static variable
			Next
			$vUserFunction_Static = $avUserFuncWParams

		Else
			$vUserFunction_Static = $vUserFunction
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)

	ElseIf $vUserFunction = Null Then
		; Clear User Function.
		$vUserFunction_Static = Default

		Return SetError($__LO_STATUS_SUCCESS, 0, 2)

	Else
		; return error as an incorrect parameter was passed to this function

		Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOCalc_ComError_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FilterDescriptorCreate
; Description ...: Create a Filter Descriptor to use in the Filter function.
; Syntax ........: _LOCalc_FilterDescriptorCreate(ByRef $oRange, $atFilterField[, $bCaseSensitive = False[, $bSkipDupl = False[, $bUseRegExp = False[, $bHeaders = False[, $bCopyOutput = False[, $oCopyOutput = Null[, $bSaveCriteria = True]]]]]]])
; Parameters ....: $oRange              - The Range you intend to apply the Filter to. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetActive function.
;                  $atFilterField       - A single column Array of Filter Fields previously created by _LOCalc_FilterFieldCreate. Maximum of 8 Fields allowed.
;                  $bCaseSensitive      - [optional] Default is False. If True, the Filtering operation will be case sensitive.
;                  $bSkipDupl           - [optional] Default is False. If True, Duplicate values will be skipped in the list of filtered data.
;                  $bUseRegExp          - [optional] Default is False. If True, the String Value set will be considered as using Regular expressions.
;                  $bHeaders            - [optional] Default is False. If True, the Range contains column headers.
;                  $bCopyOutput         - [optional] Default is False. If True, the filtering results are copied to another location in the Sheet.
;                  $oCopyOutput         - [optional] Default is Null. The location to copy filter data to. If a range is input, the first cell is used. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetActive function.
;                  $bSaveCriteria       - [optional] Default is True. If True, the output range remains linked to the source range, allowing for future re-application of the same filter to the range. Source Range must be previously defined as a Database range.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully created a Filter descriptor Object, returning its Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oRange not an Object.
;                  @Error: 1, @Extended: 2 = $atFilterField not an Array, or Array contains more than 8 elements.
;                  @Error: 1, @Extended: 3 = $bCaseSensitive not a Boolean.
;                  @Error: 1, @Extended: 4 = $bSkipDupl not a Boolean.
;                  @Error: 1, @Extended: 5 = $bUseRegExp not a Boolean.
;                  @Error: 1, @Extended: 6 = $bHeaders not a Boolean.
;                  @Error: 1, @Extended: 7 = $bCopyOutput not a Boolean.
;                  @Error: 1, @Extended: 8 = $oCopyOutput not an Object.
;                  @Error: 1, @Extended: 9 = $bSaveCriteria not a Boolean.
;                  @Error: 1, @Extended: 10 = $atFilterField contains an element that is not an Object. Returning problem element index.
;                  @Error: 1, @Extended: 11 = $bCopyOutput called with True, but $oCopyOutput not an Object.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a Filter Descriptor Object.
;                  @Error: 2, @Extended: 2 = Failed to create a "com.sun.star.table.CellAddress" Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Address for Cell or Cell Range called in $oCopyOutput.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_FilterDescriptorModify, _LOCalc_FilterFieldCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FilterDescriptorCreate(ByRef $oRange, $atFilterField, $bCaseSensitive = False, $bSkipDupl = False, $bUseRegExp = False, $bHeaders = False, $bCopyOutput = False, $oCopyOutput = Null, $bSaveCriteria = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFilterDesc
	Local $tCellInputAddr, $tCellAddr
	Local Const $__LOC_FILTER_ORIENTATION_ROWS = 1 ; Orientation isn't implemented in L.O. so Rows is the only option.

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($atFilterField) Or (UBound($atFilterField) > 8) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bCaseSensitive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bSkipDupl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bUseRegExp) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bHeaders) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If ($oCopyOutput <> Null) And Not IsObj($oCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If Not IsBool($bSaveCriteria) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	$oFilterDesc = $oRange.createFilterDescriptor(True)
	If Not IsObj($oFilterDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($atFilterField) - 1
		If Not IsObj($atFilterField[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, $i)
	Next

	If ($bCopyOutput = True) Then
		If Not IsObj($oCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tCellInputAddr = $oCopyOutput.RangeAddress()
		If Not IsObj($tCellInputAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$tCellAddr = __LO_CreateStruct("com.sun.star.table.CellAddress")
		If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$tCellAddr.Sheet = $tCellInputAddr.Sheet()
		$tCellAddr.Column = $tCellInputAddr.StartColumn()
		$tCellAddr.Row = $tCellInputAddr.StartRow()
	EndIf

	; Orientation is only set to Rows. I tried setting it to columns, but it doesn't work. Seemingly Filtering Columns isn't implemented yet, which is confirmed by a
	; post from 2009 by Villeroy on the OpenOffice Forums inside of a Macro posted.
	; https://forum.openoffice.org/en/forum/viewtopic.php?p=78786&sid=1e046304b59035364caecb0ad0a10327#p78786

	With $oFilterDesc
		.setFilterFields2($atFilterField)
		.IsCaseSensitive = $bCaseSensitive
		.SkipDuplicates = $bSkipDupl
		.UseRegularExpressions = $bUseRegExp
		.ContainsHeader = $bHeaders
		.Orientation = $__LOC_FILTER_ORIENTATION_ROWS
		.CopyOutputData = $bCopyOutput
		.SaveOutputPosition = $bSaveCriteria
	EndWith

	If IsObj($oCopyOutput) Then $oFilterDesc.OutputPosition = $tCellAddr

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFilterDesc)
EndFunc   ;==>_LOCalc_FilterDescriptorCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FilterDescriptorModify
; Description ...: Set or Retrieve Filter Descriptor settings.
; Syntax ........: _LOCalc_FilterDescriptorModify(ByRef $oRange, ByRef $oFilterDesc[, $atFilterField = Null[, $bCaseSensitive = Null[, $bSkipDupl = Null[, $bUseRegExp = Null[, $bHeaders = Null[, $bCopyOutput = Null[, $oCopyOutput = Null[, $bSaveCriteria = Null]]]]]]]])
; Parameters ....: $oRange              - The Sheet the Filter Descriptor was Created with, or the Range you intend to apply the Filter to. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetActive function.
;                  $oFilterDesc         - A Filter Descriptor created by a previous _LOCalc_FilterDescriptorCreate function.
;                  $atFilterField       - [optional] Default is Null. A single column Array of Filter Fields previously created by _LOCalc_FilterFieldCreate. Maximum of 8 Fields allowed.
;                  $bCaseSensitive      - [optional] Default is Null. If True, the Filtering operation will be case sensitive.
;                  $bSkipDupl           - [optional] Default is Null. If True, Duplicate values will be skipped in the list of filtered data.
;                  $bUseRegExp          - [optional] Default is Null. If True, the String Value set will be considered as using Regular expressions.
;                  $bHeaders            - [optional] Default is Null. If True, the Range contains column headers.
;                  $bCopyOutput         - [optional] Default is Null. If True, the filtering results are copied to another location in the Sheet.
;                  $oCopyOutput         - [optional] Default is Null. The location to copy filter data to. If a range is input, the first cell is used. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetActive function.
;                  $bSaveCriteria       - [optional] Default is Null. If True, the output range remains linked to the source range, allowing for future re-application of the same filter to the range. Source Range must be previously defined as a Database range.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Filter Descriptor was successfully modified.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oRange not an Object.
;                  @Error: 1, @Extended: 2 = $oFilterDesc not an Object.
;                  @Error: 1, @Extended: 3 = $atFilterField not an Array, or Array contains more than 8 elements.
;                  @Error: 1, @Extended: 4 = $atFilterField contains an element that is not an Object. Returning problem element index.
;                  @Error: 1, @Extended: 5 = $bCaseSensitive not a Boolean.
;                  @Error: 1, @Extended: 6 = $bSkipDupl not a Boolean.
;                  @Error: 1, @Extended: 7 = $bUseRegExp not a Boolean.
;                  @Error: 1, @Extended: 8 = $bHeaders not a Boolean.
;                  @Error: 1, @Extended: 9 = $bCopyOutput not a Boolean.
;                  @Error: 1, @Extended: 10 = $bCopyOutput called with True, but $oCopyOutput not an Object.
;                  @Error: 1, @Extended: 11 = $oCopyOutput not an Object.
;                  @Error: 1, @Extended: 12 = $bSaveCriteria not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a "com.sun.star.table.CellAddress" Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object for Cell referenced in $oCopyOutput.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Cell Address for Cell or Cell Range called in $oCopyOutput.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When retrieving the current settings for a filter descriptor, the Return value for $oCopyOutput is a single Cell Object.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_FilterDescriptorCreate, _LOCalc_FilterFieldCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FilterDescriptorModify(ByRef $oRange, ByRef $oFilterDesc, $atFilterField = Null, $bCaseSensitive = Null, $bSkipDupl = Null, $bUseRegExp = Null, $bHeaders = Null, $bCopyOutput = Null, $oCopyOutput = Null, $bSaveCriteria = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avFilter[8]
	Local $tCellInputAddr, $tCellAddr
	Local $oCell

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFilterDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($atFilterField, $bCaseSensitive, $bSkipDupl, $bUseRegExp, $bHeaders, $bCopyOutput, $oCopyOutput, $bSaveCriteria) Then
		$oCell = $oRange.Spreadsheet.getCellByPosition($oFilterDesc.OutputPosition.Column(), $oFilterDesc.OutputPosition.Row())
		If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		__LO_ArrayFill($avFilter, $oFilterDesc.getFilterFields2(), $oFilterDesc.IsCaseSensitive(), $oFilterDesc.SkipDuplicates(), $oFilterDesc.UseRegularExpressions(), _
				$oFilterDesc.ContainsHeader(), $oFilterDesc.CopyOutputData(), $oCell, $oFilterDesc.SaveOutputPosition())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFilter)
	EndIf

	If ($atFilterField <> Null) Then
		If Not IsArray($atFilterField) Or Not (UBound($atFilterField) <= 8) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		For $i = 0 To UBound($atFilterField) - 1
			If Not IsObj($atFilterField[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
		Next

		$oFilterDesc.setFilterFields2($atFilterField)
	EndIf

	If ($bCaseSensitive <> Null) Then
		If Not IsBool($bCaseSensitive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oFilterDesc.IsCaseSensitive = $bCaseSensitive
	EndIf

	If ($bSkipDupl <> Null) Then
		If Not IsBool($bSkipDupl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oFilterDesc.SkipDuplicates = $bSkipDupl
	EndIf

	If ($bUseRegExp <> Null) Then
		If Not IsBool($bUseRegExp) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oFilterDesc.UseRegularExpressions = $bUseRegExp
	EndIf

	If ($bHeaders <> Null) Then
		If Not IsBool($bHeaders) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oFilterDesc.ContainsHeader = $bHeaders
	EndIf

	If ($bCopyOutput <> Null) Then
		If Not IsBool($bCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If ($bCopyOutput = True) And Not IsObj($oCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oFilterDesc.CopyOutputData = $bCopyOutput
	EndIf

	If ($oCopyOutput <> Null) Then
		If Not IsObj($oCopyOutput) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tCellInputAddr = $oCopyOutput.RangeAddress()
		If Not IsObj($tCellInputAddr) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$tCellAddr = __LO_CreateStruct("com.sun.star.table.CellAddress")
		If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tCellAddr.Sheet = $tCellInputAddr.Sheet()
		$tCellAddr.Column = $tCellInputAddr.StartColumn()
		$tCellAddr.Row = $tCellInputAddr.StartRow()

		$oFilterDesc.OutputPosition = $tCellAddr
	EndIf

	If ($bSaveCriteria <> Null) Then
		If Not IsBool($bSaveCriteria) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oFilterDesc.SaveOutputPosition = $bSaveCriteria
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_FilterDescriptorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FilterFieldCreate
; Description ...: Create a Filter Field for defining Filter values and settings.
; Syntax ........: _LOCalc_FilterFieldCreate($iColumn[, $bIsNumeric = False[, $nValue = 0[, $sString = ""[, $iCondition = $LOC_FILTER_CONDITION_EMPTY[, $iOperator = $LOC_FILTER_OPERATOR_AND]]]]])
; Parameters ....: $iColumn             - The 0 based Column number to perform the filtering operation upon counting from the beginning of the range.
;                  $bIsNumeric          - [optional] Default is False. If True, the filter Value to search for is a number. If False, the filter value to search for is a string.
;                  $nValue              - [optional] Default is 0. The numerical Value to filter the Range for. Only valid if $bIsNumeric is set to True. Call with any number to skip, it will not be used unless $bIsNumeric is True.
;                  $sString             - [optional] Default is "". The string Value to filter the Range for. Only valid if $bIsNumeric is set to False. Call with an empty string to skip, it will not be used unless $bIsNumeric is False.
;                  $iCondition          - [optional] (0-17) Default is $LOC_FILTER_CONDITION_EMPTY. The comparative condition to test each cell and value by. See Constants $LOC_FILTER_CONDITION_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iOperator           - [optional] (0, 1) Default is $LOC_FILTER_OPERATOR_AND. The connection this filter field has with the previous filter field. See Constants $LOC_FILTER_OPERATOR_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Struct
;                  @Error: 0, @Extended: 0, Return: Struct = Success. Successfully created and returned the Filter Field Structure.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $iColumn not an Integer.
;                  @Error: 1, @Extended: 2 = $bIsNumeric not a Boolean.
;                  @Error: 1, @Extended: 3 = $nValue not a number.
;                  @Error: 1, @Extended: 4 = $sString not a String.
;                  @Error: 1, @Extended: 5 = $iCondition not an Integer, less than 0 or greater than 17. See Constants $LOC_FILTER_CONDITION_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iOperator not an Integer, less than 0 or greater than 1. See Constants $LOC_FILTER_OPERATOR_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a "com.sun.star.sheet.TableFilterField2" Struct.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Filter Descriptor can contain up to 8 of these Filter Fields. Once you create the Filter Field Structure, place it in an array before using it to create a Filter descriptor. Place each Filter Field Structure in a separate element of the Array.
; Related .......: _LOCalc_FilterFieldModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FilterFieldCreate($iColumn, $bIsNumeric = False, $nValue = 0, $sString = "", $iCondition = $LOC_FILTER_CONDITION_EMPTY, $iOperator = $LOC_FILTER_OPERATOR_AND)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tFilterField

	If Not IsInt($iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bIsNumeric) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iCondition, $LOC_FILTER_CONDITION_EMPTY, $LOC_FILTER_CONDITION_DOES_NOT_END_WITH) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not __LO_IntIsBetween($iOperator, $LOC_FILTER_OPERATOR_AND, $LOC_FILTER_OPERATOR_OR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$tFilterField = __LO_CreateStruct("com.sun.star.sheet.TableFilterField2")
	If Not IsObj($tFilterField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	With $tFilterField
		.Field = $iColumn
		.IsNumeric = $bIsNumeric
		.NumericValue = $nValue
		.StringValue = $sString
		.Operator = $iCondition ; L.O. calls Operator "Condition" in U.I.
		.Connection = $iOperator ; L.O. calls Connection "Operator" in U.I.
	EndWith

	Return SetError($__LO_STATUS_SUCCESS, 0, $tFilterField)
EndFunc   ;==>_LOCalc_FilterFieldCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FilterFieldModify
; Description ...: Set or Retrieve Filter Field structure settings.
; Syntax ........: _LOCalc_FilterFieldModify(ByRef $tFilterField[, $iColumn = Null[, $bIsNumeric = Null[, $nValue = Null[, $sString = Null[, $iCondition = Null[, $iOperator = Null]]]]]])
; Parameters ....: $tFilterField        - A Filter Field from a previous _LOCalc_FilterFieldCreate function call.
;                  $iColumn             - [optional] Default is Null. The 0 based Column number to perform the filtering operation upon counting from the beginning of the range.
;                  $bIsNumeric          - [optional] Default is Null. If True, the filter Value to search for is a number. If False, the filter value to search for is a string.
;                  $nValue              - [optional] Default is Null. The numerical Value to filter the Range for. Only valid if $bIsNumeric is set to True.
;                  $sString             - [optional] Default is Null. The string Value to filter the Range for. Only valid if $bIsNumeric is set to False.
;                  $iCondition          - [optional] (0-17) Default is Null. The comparative condition to test each cell and value by. See Constants $LOC_FILTER_CONDITION_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iOperator           - [optional] (0, 1) Default is Null. The connection this filter field has with the previous filter field. See Constants $LOC_FILTER_OPERATOR_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Struct
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Filter Field Structure was successfully modified.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $tFilterField not an Object.
;                  @Error: 1, @Extended: 2 = $iColumn not an Integer.
;                  @Error: 1, @Extended: 3 = $bIsNumeric not a Boolean.
;                  @Error: 1, @Extended: 4 = $nValue not a number.
;                  @Error: 1, @Extended: 5 = $sString not a String.
;                  @Error: 1, @Extended: 6 = $iCondition not an Integer, less than 0 or greater than 17. See Constants $LOC_FILTER_CONDITION_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error: 1, @Extended: 7 = $iOperator not an Integer, less than 0 or greater than 1. See Constants $LOC_FILTER_OPERATOR_* as defined in LibreOfficeCalc_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Filter Descriptor can contain up to 8 of these Filter Fields. Once you create the Filter Field Structure, place it in an array before using it to create a Filter descriptor. Place each Filter Field Structure in a separate element of the Array.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_FilterFieldCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FilterFieldModify(ByRef $tFilterField, $iColumn = Null, $bIsNumeric = Null, $nValue = Null, $sString = Null, $iCondition = Null, $iOperator = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avFilter[6]

	If Not IsObj($tFilterField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iColumn, $bIsNumeric, $nValue, $sString, $iCondition, $iOperator) Then
		__LO_ArrayFill($avFilter, $tFilterField.Field(), $tFilterField.IsNumeric(), $tFilterField.NumericValue(), $tFilterField.StringValue(), $tFilterField.Operator(), $tFilterField.Connection())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFilter)
	EndIf

	If ($iColumn <> Null) Then
		If Not IsInt($iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tFilterField.Field = $iColumn
	EndIf

	If ($bIsNumeric <> Null) Then
		If Not IsBool($bIsNumeric) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tFilterField.IsNumeric = $bIsNumeric
	EndIf

	If ($nValue <> Null) Then
		If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tFilterField.NumericValue = $nValue
	EndIf

	If ($sString <> Null) Then
		If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tFilterField.StringValue = $sString
	EndIf

	If ($iCondition <> Null) Then ; L.O. calls Operator "Condition" in U.I.
		If Not __LO_IntIsBetween($iCondition, $LOC_FILTER_CONDITION_EMPTY, $LOC_FILTER_CONDITION_DOES_NOT_END_WITH) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tFilterField.Operator = $iCondition
	EndIf

	If ($iOperator <> Null) Then ; L.O. calls Connection "Operator" in U.I.
		If Not __LO_IntIsBetween($iOperator, $LOC_FILTER_OPERATOR_AND, $LOC_FILTER_OPERATOR_OR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tFilterField.Connection = $iOperator
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_FilterFieldModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FontExists
; Description ...: Tests whether a specific font exists by name.
; Syntax ........: _LOCalc_FontExists($sFontName[, $oDoc = Null])
; Parameters ....: $sFontName           - The Font name to search for.
;                  $oDoc                - [optional] Default is Null. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Font is available, else False.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $sFontName not a String.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  @Error: 2, @Extended: 3 = Failed to create a Property Struct.
;                  @Error: 2, @Extended: 4 = Failed to create a new Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Font list.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oDoc is optional, if not called, a Calc Document is created invisibly to perform the check.
; Related .......: _LOCalc_FontsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FontExists($sFontName, $oDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts, $atProperties[1]
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $oServiceManager, $oDesktop
	Local $bClose = False

	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If Not IsObj($oDoc) Then
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
		If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$atProperties[0] = __LO_SetPropertyValue("Hidden", True)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$oDoc = $oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", $iURLFrameCreate, $atProperties)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

		$bClose = True
	EndIf

	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then
		If $bClose Then $oDoc.Close(True)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	For $i = 0 To UBound($atFonts) - 1
		If $atFonts[$i].Name = $sFontName Then
			If $bClose Then $oDoc.Close(True)

			Return SetError($__LO_STATUS_SUCCESS, 0, True)
		EndIf
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If $bClose Then $oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOCalc_FontExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FontsGetNames
; Description ...: Retrieve an array of currently available font names.
; Syntax ........: _LOCalc_FontsGetNames([$oDoc = Null])
; Parameters ....: $oDoc                - [optional] Default is Null. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning a 4 Column Array, @Extended is set to the number of results. See remarks
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  @Error: 2, @Extended: 3 = Failed to create a Property Struct.
;                  @Error: 2, @Extended: 4 = Failed to create a new Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Font list.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oDoc is optional, if not called, a Calc Document is created invisibly to perform the check.
;                  Many fonts will be listed multiple times, this is because of the varying settings for them, such as bold, Italic, etc. Style Name is really a repeat of weight(Bold) and Slant (Italic) settings, but is included for easier processing if required.
;                  From personal tests, Slant only returns 0 or 2.
;                  The returned array will be as follows:
;                  The first column (Array[1][0]) contains the Font Name.
;                  The Second column (Array [1][1] contains the style name (Such as Bold Italic etc.)
;                  The third column (Array[1][2]) contains the Font weight (Bold) See Constants, $LOC_CHAR_WEIGHT_* as defined in LibreOfficeCalc_Constants.au3;
;                  The fourth column (Array[1][3]) contains the font slant (Italic) See constants, $LOC_CHAR_POSTURE_* as defined in LibreOfficeCalc_Constants.au3.
; Related .......: _LOCalc_FontExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FontsGetNames($oDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asFonts[0][4]
	Local $atFonts, $atProperties[1]
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $oServiceManager, $oDesktop
	Local $bClose = False

	If Not IsObj($oDoc) Then
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
		If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$atProperties[0] = __LO_SetPropertyValue("Hidden", True)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$oDoc = $oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", $iURLFrameCreate, $atProperties)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

		$bClose = True
	EndIf

	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then
		If $bClose Then $oDoc.Close(True)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	ReDim $asFonts[UBound($atFonts)][4]

	For $i = 0 To UBound($atFonts) - 1
		$asFonts[$i][0] = $atFonts[$i].Name()
		$asFonts[$i][1] = $atFonts[$i].StyleName()
		$asFonts[$i][2] = $atFonts[$i].Weight
		$asFonts[$i][3] = $atFonts[$i].Slant() ; only 0 or 2?
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If $bClose Then $oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, UBound($atFonts), $asFonts)
EndFunc   ;==>_LOCalc_FontsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeyCreate
; Description ...: Create a Format Key.
; Syntax ........: _LOCalc_FormatKeyCreate(ByRef $oDoc, $sFormat)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sFormat             - The format key String to create.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Format Key was successfully created, returning Format Key Integer.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. Format Key already existed, returning Format Key Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sFormat not a String.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to Create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Number Formats Object.
;                  @Error: 3, @Extended: 2 = Failed to Create or Retrieve the Format key.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_FormatKeyDelete, _LOCalc_FormatKeyGetStandard
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeyCreate(ByRef $oDoc, $sFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFormatKey
	Local $tLocale
	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tLocale = __LO_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iFormatKey = $oFormats.queryKey($sFormat, $tLocale, False)
	If ($iFormatKey > -1) Then Return SetError($__LO_STATUS_SUCCESS, 1, $iFormatKey) ; Format already existed
	$iFormatKey = $oFormats.addNew($sFormat, $tLocale)
	If ($iFormatKey > -1) Then Return SetError($__LO_STATUS_SUCCESS, 0, $iFormatKey) ; Format created

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Failed to create or retrieve Format
EndFunc   ;==>_LOCalc_FormatKeyCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeyDelete
; Description ...: Delete a User-Created Format Key from a Document.
; Syntax ........: _LOCalc_FormatKeyDelete(ByRef $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iFormatKey          - The User-Created format Key to delete.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Format Key was successfully deleted.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iFormatKey not an Integer.
;                  @Error: 1, @Extended: 3 = Format Key called in $iFormatKey not found in Document.
;                  @Error: 1, @Extended: 4 = Format Key called in $iFormatKey not User-Created.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Number Formats Object.
;                  @Error: 3, @Extended: 2 = Failed to delete key.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_FormatKeysGetList, _LOCalc_FormatKeyCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeyDelete(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOCalc_FormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Key not found.

	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($oFormats.getbykey($iFormatKey).UserDefined() = False) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Key not User Created.

	$oFormats.removeByKey($iFormatKey)
	If _LOCalc_FormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_FormatKeyDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeyExists
; Description ...: Check if a Document contains a certain Format Key.
; Syntax ........: _LOCalc_FormatKeyExists(ByRef $oDoc, $iFormatKey[, $iFormatType = $LOC_FORMAT_KEYS_ALL])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iFormatKey          - The Format Key to look for.
;                  $iFormatType         - [optional] (0-15881) Default is $LOC_FORMAT_KEYS_ALL. The Format Key type to search in. Values can be BitOr'd together. See Constants, $LOC_FORMAT_KEYS_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. If the Format Key exists in document, True is returned, else False.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iFormatKey not an Integer.
;                  @Error: 1, @Extended: 3 = $iFormatType not an Integer, less than 0 or greater than 15881. See Constants, $LOC_FORMAT_KEYS_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to Create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Number Formats Object.
;                  @Error: 3, @Extended: 2 = Failed to obtain Array of Date/Time Formats.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeyExists(ByRef $oDoc, $iFormatKey, $iFormatType = $LOC_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys[0]
	Local $tLocale

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iFormatType, $LOC_FORMAT_KEYS_ALL, 15881) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; 15881 = All keys BitOR'd together.

	$tLocale = __LO_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aiFormatKeys = $oFormats.queryKeys($iFormatType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aiFormatKeys) - 1
		If ($aiFormatKeys[$i] = $iFormatKey) Then Return SetError($__LO_STATUS_SUCCESS, 0, True) ; Doc does contain format Key
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, False) ; Doc does not contain format Key
EndFunc   ;==>_LOCalc_FormatKeyExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeyGetStandard
; Description ...: Retrieve the Standard Format for a specific Format Key Type.
; Syntax ........: _LOCalc_FormatKeyGetStandard(ByRef $oDoc, $iFormatKeyType)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iFormatKeyType      - (1-8196) The Format Key type to retrieve the standard Format for. See Constants $LOC_FORMAT_KEYS_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning the Standard Format for the requested Format Key Type.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iFormatKeyType not an Integer, less than 1 or greater than 8196. See Constants $LOC_FORMAT_KEYS_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a "com.sun.star.lang.Locale" Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Number Formats Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve the Standard Format for the requested Format Key Type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeyGetStandard(ByRef $oDoc, $iFormatKeyType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $tLocale
	Local $iStandard

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iFormatKeyType, $LOC_FORMAT_KEYS_DEFINED, $LOC_FORMAT_KEYS_DURATION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tLocale = __LO_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iStandard = $oFormats.getStandardFormat($iFormatKeyType, $tLocale)
	If Not IsInt($iStandard) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iStandard)
EndFunc   ;==>_LOCalc_FormatKeyGetStandard

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeyGetString
; Description ...: Retrieve a Format Key String.
; Syntax ........: _LOCalc_FormatKeyGetString(ByRef $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $iFormatKey          - The Format Key to retrieve the string for.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Success. Returning Format Key's Format String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iFormatKey not an Integer.
;                  @Error: 1, @Extended: 3 = $iFormatKey not found in Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve requested Format Key Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Format Key String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_FormatKeysGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeyGetString(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormatKey
	Local $sFormatkKey

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOCalc_FormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oFormatKey = $oDoc.getNumberFormats().getByKey($iFormatKey)
	If Not IsObj($oFormatKey) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Key not found.

	$sFormatkKey = $oFormatKey.FormatString()
	If Not IsString($sFormatkKey) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sFormatkKey)
EndFunc   ;==>_LOCalc_FormatKeyGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_FormatKeysGetList
; Description ...: Retrieve an Array of Date/Time Format Keys.
; Syntax ........: _LOCalc_FormatKeysGetList(ByRef $oDoc[, $bIsUser = False[, $bUserOnly = False[, $iFormatKeyType = $LOC_FORMAT_KEYS_ALL]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $bIsUser             - [optional] Default is False. If True, Adds a third column to the return Array with a boolean, whether each Key is user-created or not.
;                  $bUserOnly           - [optional] Default is False. If True, only user-created Format Keys are returned.
;                  $iFormatKeyType      - [optional] (0-15881) Default is $LOC_FORMAT_KEYS_ALL. The Format Key type to retrieve an array of. Values can be BitOr'd together. See Constants, $LOC_FORMAT_KEYS_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning a 2 or 3 column Array, depending on current $bIsUser setting. See remarks. @Extended is set to the number of Keys returned.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bIsUser not a Boolean.
;                  @Error: 1, @Extended: 3 = $bUserOnly not a Boolean.
;                  @Error: 1, @Extended: 4 = $iFormatKeyType not an Integer, less than 0 or greater than 15881. See Constants, $LOC_FORMAT_KEYS_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve NumberFormats Object.
;                  @Error: 3, @Extended: 2 = Failed to obtain Array of Format Keys.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Column One (Array[0][0]) will contain the Format Key Integer,
;                  Column two (Array[0][1]) will contain the Format Key String,
;                  If $bIsUser is called with True, Column Three (Array[0][2]) will contain a Boolean, True if the Format Key is User-created, else False.
; Related .......: _LOCalc_FormatKeyDelete, _LOCalc_FormatKeyGetString, _LOCalc_FormatKeyGetStandard
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_FormatKeysGetList(ByRef $oDoc, $bIsUser = False, $bUserOnly = False, $iFormatKeyType = $LOC_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys
	Local $avFormats[0][3]
	Local $tLocale
	Local $iColumns = 3, $iCount = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bIsUser) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iColumns = ($bIsUser = True) ? ($iColumns) : (2)

	If Not __LO_IntIsBetween($iFormatKeyType, $LOC_FORMAT_KEYS_ALL, 15881) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; 15881 = all keys BitOR'd together.

	$tLocale = __LO_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aiFormatKeys = $oFormats.queryKeys($iFormatKeyType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	ReDim $avFormats[UBound($aiFormatKeys)][$iColumns]

	For $i = 0 To UBound($aiFormatKeys) - 1
		If ($bUserOnly = True) Then
			If ($oFormats.getbykey($aiFormatKeys[$i]).UserDefined() = True) Then
				$avFormats[$iCount][0] = $aiFormatKeys[$i]
				$avFormats[$iCount][1] = $oFormats.getbykey($aiFormatKeys[$i]).FormatString()
				If ($bIsUser = True) Then $avFormats[$iCount][2] = $oFormats.getbykey($aiFormatKeys[$i]).UserDefined()
				$iCount += 1
			EndIf

		Else
			$avFormats[$i][0] = $aiFormatKeys[$i]
			$avFormats[$i][1] = $oFormats.getbykey($aiFormatKeys[$i]).FormatString()
			If ($bIsUser = True) Then $avFormats[$i][2] = $oFormats.getbykey($aiFormatKeys[$i]).UserDefined()
		EndIf
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	If ($bUserOnly = True) Then ReDim $avFormats[$iCount][$iColumns]

	Return SetError($__LO_STATUS_SUCCESS, UBound($avFormats), $avFormats)
EndFunc   ;==>_LOCalc_FormatKeysGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SearchDescriptorCreate
; Description ...: Create a Search Descriptor for searching a document.
; Syntax ........: _LOCalc_SearchDescriptorCreate(ByRef $oRange[, $bBackwards = False[, $bSearchRows = True[, $bMatchCase = False[, $iSearchIn = $LOC_SEARCH_IN_FORMULAS[, $bEntireCell = False[, $bRegExp = False[, $bWildcards = False[, $bStyles = False]]]]]]]])
; Parameters ....: $oRange              - A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetActive function.
;                  $bBackwards          - [optional] Default is False. If True, the document is searched backwards.
;                  $bSearchRows         - [optional] Default is True. If True, Search is performed left to right along the rows, else if False, the search is performed top to bottom along the columns.
;                  $bMatchCase          - [optional] Default is False. If True, the case of the letters is important for the Search.
;                  $iSearchIn           - [optional] (0-2) Default is $LOC_SEARCH_IN_FORMULAS. The Cell data type to search in. See Constants $LOC_SEARCH_IN_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bEntireCell         - [optional] Default is False. If True, Searches for whole words or cells that are identical to the search text.
;                  $bRegExp             - [optional] Default is False. If True, the search string is evaluated as a regular expression.
;                  $bWildcards          - [optional] Default is False. If True, the search string is considered to contain wildcards (* ?). A Backslash can be used to escape a wildcard.
;                  $bStyles             - [optional] Default is False. If True, the search string is considered a Cell Style name, and the search will return any Cell utilizing the specified name.
; Return values .: Success: Object.
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning a Search Descriptor Object for setting Search options.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oRange not an Object.
;                  @Error: 1, @Extended: 2 = $bBackwards not a Boolean.
;                  @Error: 1, @Extended: 3 = $bSearchRows not a Boolean.
;                  @Error: 1, @Extended: 4 = $bMatchCase not a Boolean.
;                  @Error: 1, @Extended: 5 = $iSearchIn not an Integer, less than 0 or greater than 2. See Constants $LOC_SEARCH_IN_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error: 1, @Extended: 6 = $bEntireCell not a Boolean.
;                  @Error: 1, @Extended: 7 = $bRegExp not a Boolean.
;                  @Error: 1, @Extended: 8 = $bWildcards not a Boolean.
;                  @Error: 1, @Extended: 9 = $bStyles not a Boolean.
;                  @Error: 1, @Extended: 10 = Both $bRegExp and $bWildcards are called with True, only one can be True at one time.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create Search Descriptor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The returned Search Descriptor is only good for the Document that contained the Range it was created by, it WILL NOT work for other Documents.
; Related .......: _LOCalc_SearchDescriptorModify, _LOCalc_SearchDescriptorSimilarityModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SearchDescriptorCreate(ByRef $oRange, $bBackwards = False, $bSearchRows = True, $bMatchCase = False, $iSearchIn = $LOC_SEARCH_IN_FORMULAS, $bEntireCell = False, $bRegExp = False, $bWildcards = False, $bStyles = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSrchDescript

	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bBackwards) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bSearchRows) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bMatchCase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iSearchIn, $LOC_SEARCH_IN_FORMULAS, $LOC_SEARCH_IN_COMMENTS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bEntireCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bRegExp) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not IsBool($bWildcards) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If Not IsBool($bStyles) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If ($bWildcards = True) And ($bRegExp = True) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

	$oSrchDescript = $oRange.createSearchDescriptor()
	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	With $oSrchDescript
		.SearchBackwards = $bBackwards
		.SearchByRow = $bSearchRows
		.SearchCaseSensitive = $bMatchCase
		.SearchType = $iSearchIn
		.SearchWords = $bEntireCell
		.SearchWildcard = $bWildcards
		; Regular Expression setting MUST be after Wildcards, setting Wildcards to False, even when it is already set to False, changes RegExp to False no matter what.
		; -- Slated to be fixed L.O. 24.8.0
		.SearchRegularExpression = $bRegExp
		.SearchStyles = $bStyles
	EndWith

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSrchDescript)
EndFunc   ;==>_LOCalc_SearchDescriptorCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SearchDescriptorModify
; Description ...: Modify Search Descriptor settings of an existing Search Descriptor Object.
; Syntax ........: _LOCalc_SearchDescriptorModify(ByRef $oSrchDescript[, $bBackwards = Null[, $bSearchRows = Null[, $bMatchCase = Null[, $iSearchIn = Null[, $bEntireCell = Null[, $bRegExp = Null[, $bWildcards = Null[, $bStyles = Null]]]]]]]])
; Parameters ....: $oSrchDescript       - A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $bBackwards          - [optional] Default is Null. If True, the document is searched backwards.
;                  $bSearchRows         - [optional] Default is Null. If True, Search is performed left to right along the rows, else if False, the search is performed top to bottom along the columns.
;                  $bMatchCase          - [optional] Default is Null. If True, the case of the letters is important for the Search.
;                  $iSearchIn           - [optional] (0-2) Default is Null. The Cell data type to search in. See Constants $LOC_SEARCH_IN_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bEntireCell         - [optional] Default is Null. If True, Searches for whole words or cells that are identical to the search text.
;                  $bRegExp             - [optional] Default is Null. If True, the search string is evaluated as a regular expression.
;                  $bWildcards          - [optional] Default is Null. If True, the search string is considered to contain wildcards (* ?). A Backslash can be used to escape a wildcard.
;                  $bStyles             - [optional] Default is Null. If True, the search string is considered a Cell Style name, and the search will return any Cell utilizing the specified name.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Returning 1 after directly modifying Search Descriptor Object.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 8 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSrchDescript not an Object.
;                  @Error: 1, @Extended: 2 = $oSrchDescript Object not a Search Descriptor Object.
;                  @Error: 1, @Extended: 3 = $bBackwards not a Boolean.
;                  @Error: 1, @Extended: 4 = $bSearchRows not a Boolean.
;                  @Error: 1, @Extended: 5 = $bMatchCase not a Boolean.
;                  @Error: 1, @Extended: 6 = $iSearchIn not an Integer, less than 0 or greater than 2. See Constants $LOC_SEARCH_IN_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error: 1, @Extended: 7 = $bEntireCell not a Boolean.
;                  @Error: 1, @Extended: 8 = $bRegExp not a Boolean.
;                  @Error: 1, @Extended: 9 = $bWildcards not a Boolean.
;                  @Error: 1, @Extended: 10 = $bStyles not a Boolean.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When calling $bRegExp or $bWildcards with True, if any of following three are set to True, they will be set to False: $bSimilarity(From the Similarity function), $bRegExp or $bWildcards.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_SearchDescriptorCreate, _LOCalc_SearchDescriptorSimilarityModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SearchDescriptorModify(ByRef $oSrchDescript, $bBackwards = Null, $bSearchRows = Null, $bMatchCase = Null, $iSearchIn = Null, $bEntireCell = Null, $bRegExp = Null, $bWildcards = Null, $bStyles = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSrchDescript[8]

	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bBackwards, $bSearchRows, $bMatchCase, $iSearchIn, $bEntireCell, $bRegExp, $bWildcards, $bStyles) Then
		__LO_ArrayFill($avSrchDescript, $oSrchDescript.SearchBackwards(), $oSrchDescript.SearchByRow(), $oSrchDescript.SearchCaseSensitive(), _
				$oSrchDescript.SearchType(), $oSrchDescript.SearchWords(), $oSrchDescript.SearchRegularExpression(), $oSrchDescript.SearchWildcard(), _
				$oSrchDescript.SearchStyles())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSrchDescript)
	EndIf

	If ($bBackwards <> Null) Then
		If Not IsBool($bBackwards) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oSrchDescript.SearchBackwards = $bBackwards
	EndIf

	If ($bSearchRows <> Null) Then
		If Not IsBool($bSearchRows) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oSrchDescript.SearchByRow = $bSearchRows
	EndIf

	If ($bMatchCase <> Null) Then
		If Not IsBool($bMatchCase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oSrchDescript.SearchCaseSensitive = $bMatchCase
	EndIf

	If ($iSearchIn <> Null) Then
		If Not __LO_IntIsBetween($iSearchIn, $LOC_SEARCH_IN_FORMULAS, $LOC_SEARCH_IN_COMMENTS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oSrchDescript.SearchType = $iSearchIn
	EndIf

	If ($bEntireCell <> Null) Then
		If Not IsBool($bEntireCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oSrchDescript.SearchWords = $bEntireCell
	EndIf

	If ($bWildcards <> Null) Then
		If Not IsBool($bWildcards) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		If ($bWildcards = True) And ($oSrchDescript.SearchSimilarity = True) Then $oSrchDescript.SearchSimilarity = False
		If ($bWildcards = True) And ($oSrchDescript.SearchRegularExpression = True) Then $oSrchDescript.SearchRegularExpression = False
		$oSrchDescript.SearchWildcard = $bWildcards
	EndIf
	; Regular Expression setting MUST be after Wildcards, setting Wildcards to False, even when it is already set to False, changes RegExp to False no matter what.
	; -- Slated to be fixed L.O. 24.8.0
	If ($bRegExp <> Null) Then
		If Not IsBool($bRegExp) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		If ($bRegExp = True) And ($oSrchDescript.SearchSimilarity = True) Then $oSrchDescript.SearchSimilarity = False
		$oSrchDescript.SearchRegularExpression = $bRegExp
	EndIf

	If ($bStyles <> Null) Then
		If Not IsBool($bStyles) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$oSrchDescript.SearchStyles = $bStyles
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SearchDescriptorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SearchDescriptorSimilarityModify
; Description ...: Modify Similarity Search Settings for an existing Search Descriptor Object.
; Syntax ........: _LOCalc_SearchDescriptorSimilarityModify(ByRef $oSrchDescript[, $bSimilarity = Null[, $bCombine = Null[, $iRemove = Null[, $iAdd = Null[, $iExchange = Null]]]]])
; Parameters ....: $oSrchDescript       - A Search Descriptor Object returned from _LOCalc_SearchDescriptorCreate function.
;                  $bSimilarity         - [optional] Default is Null. If True, a "similarity search" is performed.
;                  $bCombine            - [optional] Default is Null. If True, all similarity rules ($iRemove, $iAdd, and $iExchange) are applied together.
;                  $iRemove             - [optional] Default is Null. Specifies the number of characters that may be ignored to match the search pattern.
;                  $iAdd                - [optional] Default is Null. Specifies the number of characters that must be added to match the search pattern.
;                  $iExchange           - [optional] Default is Null. Specifies the number of characters that must be replaced to match the search pattern.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Returning 1 after directly modifying Search Descriptor Object.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSrchDescript not an Object.
;                  @Error: 1, @Extended: 2 = $oSrchDescript Object not a Search Descriptor Object.
;                  @Error: 1, @Extended: 3 = $bSimilarity not a Boolean.
;                  @Error: 1, @Extended: 4 = $bCombine not a Boolean.
;                  @Error: 1, @Extended: 5 = $iRemove, $iAdd, or $iExchange set to a value, but $bSimilarity not called with True.
;                  @Error: 1, @Extended: 6 = $iRemove not an Integer.
;                  @Error: 1, @Extended: 7 = $iAdd not an Integer.
;                  @Error: 1, @Extended: 8 = $iExchange not an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  If $bSimilarity is called with True while Regular Expression, or Wildcards setting is set to True, those settings will be set to False.
; Related .......: _LOCalc_SearchDescriptorCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SearchDescriptorSimilarityModify(ByRef $oSrchDescript, $bSimilarity = Null, $bCombine = Null, $iRemove = Null, $iAdd = Null, $iExchange = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSrchDescript[5]

	If Not IsObj($oSrchDescript) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bSimilarity, $bCombine, $iRemove, $iAdd, $iExchange) Then
		__LO_ArrayFill($avSrchDescript, $oSrchDescript.SearchSimilarity(), $oSrchDescript.SearchSimilarityRelax(), _
				$oSrchDescript.SearchSimilarityRemove(), $oSrchDescript.SearchSimilarityAdd(), $oSrchDescript.SearchSimilarityExchange())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSrchDescript)
	EndIf

	If ($bSimilarity <> Null) Then
		If Not IsBool($bSimilarity) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		If ($bSimilarity = True) And ($oSrchDescript.SearchRegularExpression = True) Then $oSrchDescript.SearchRegularExpression = False
		If ($bSimilarity = True) And ($oSrchDescript.SearchWildcard = True) Then $oSrchDescript.SearchWildcard = False
		$oSrchDescript.SearchSimilarity = $bSimilarity
	EndIf

	If ($bCombine <> Null) Then
		If Not IsBool($bCombine) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oSrchDescript.SearchSimilarityRelax = $bCombine
	EndIf

	If Not __LO_VarsAreNull($iRemove, $iAdd, $iExchange) Then
		If ($oSrchDescript.SearchSimilarity() = False) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		If ($iRemove <> Null) Then
			If Not IsInt($iRemove) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oSrchDescript.SearchSimilarityRemove = $iRemove
		EndIf

		If ($iAdd <> Null) Then
			If Not IsInt($iAdd) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oSrchDescript.SearchSimilarityAdd = $iAdd
		EndIf

		If ($iExchange <> Null) Then
			If Not IsInt($iExchange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oSrchDescript.SearchSimilarityExchange = $iExchange
		EndIf
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SearchDescriptorSimilarityModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SortFieldCreate
; Description ...: Create a Sort Field for sorting a Range of data with.
; Syntax ........: _LOCalc_SortFieldCreate($iIndex[, $iDataType = $LOC_SORT_DATA_TYPE_AUTO[, $bAscending = True[, $bCaseSensitive = False]]])
; Parameters ....: $iIndex              - The Column or Row to perform the sort upon. 0 Based. 0 is the first Column/Row in the Cell Range.
;                  $iDataType           - [optional] (0-2) Default is $LOC_SORT_DATA_TYPE_AUTO. The type of data that will be sorted. See Constants $LOC_SORT_DATA_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;                  $bAscending          - [optional] Default is True. If True, data will be sorted into ascending order.
;                  $bCaseSensitive      - [optional] Default is False. If True, sort will be case sensitive.
; Return values .: Success: Struct
;                  @Error: 0, @Extended: 0, Return: Struct = Success. Successfully created and returned a Sort Field Struct.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $iIndex not an Integer, or less than 0.
;                  @Error: 1, @Extended: 2 = $iDataType not an Integer, less than 0 or greater than 2. See Constants $LOC_SORT_DATA_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;                  @Error: 1, @Extended: 3 = $bAscending not a Boolean.
;                  @Error: 1, @Extended: 4 = $bCaseSensitive not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a "com.sun.star.table.TableSortField" Struct.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SortFieldCreate($iIndex, $iDataType = $LOC_SORT_DATA_TYPE_AUTO, $bAscending = True, $bCaseSensitive = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tSortField

	If Not __LO_IntIsBetween($iIndex, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iDataType, $LOC_SORT_DATA_TYPE_AUTO, $LOC_SORT_DATA_TYPE_ALPHANUMERIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAscending) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bCaseSensitive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$tSortField = __LO_CreateStruct("com.sun.star.table.TableSortField")
	If Not IsObj($tSortField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	With $tSortField
		.Field = $iIndex
		.FieldType = $iDataType
		.IsAscending = $bAscending
		.IsCaseSensitive = $bCaseSensitive
	EndWith

	Return SetError($__LO_STATUS_SUCCESS, 0, $tSortField)
EndFunc   ;==>_LOCalc_SortFieldCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SortFieldModify
; Description ...: Modify or retrieve the settings for a Sort Field previously created by _LOCalc_SortFieldCreate.
; Syntax ........: _LOCalc_SortFieldModify(ByRef $tSortField[, $iIndex = Null[, $iDataType = Null[, $bAscending = Null[, $bCaseSensitive = Null]]]])
; Parameters ....: $tSortField          - A Sort Field Struct created by a previous _LOCalc_SortFieldCreate function.
;                  $iIndex              - [optional] Default is Null. The Column or Row to perform the sort upon. 0 Based. 0 is the first Column/Row in the Cell Range.
;                  $iDataType           - [optional] (0-2) Default is Null. The type of data that will be sorted. See Constants $LOC_SORT_DATA_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;                  $bAscending          - [optional] Default is Null. If True, data will be sorted into ascending order.
;                  $bCaseSensitive      - [optional] Default is Null. If True, sort will be case sensitive.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $tSortField not an Object.
;                  @Error: 1, @Extended: 2 = $iIndex not an Integer, or less than 0.
;                  @Error: 1, @Extended: 3 = $iDataType not an Integer, less than 0 or greater than 2. See Constants $LOC_SORT_DATA_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;                  @Error: 1, @Extended: 4 = $bAscending not a Boolean.
;                  @Error: 1, @Extended: 5 = $bCaseSensitive not a Boolean.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SortFieldModify(ByRef $tSortField, $iIndex = Null, $iDataType = Null, $bAscending = Null, $bCaseSensitive = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSort[4]

	If Not IsObj($tSortField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iIndex, $iDataType, $bAscending, $bCaseSensitive) Then
		__LO_ArrayFill($avSort, $tSortField.Field(), $tSortField.FieldType(), $tSortField.IsAscending(), $tSortField.IsCaseSensitive())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSort)
	EndIf

	If ($iIndex <> Null) Then
		If Not __LO_IntIsBetween($iIndex, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tSortField.Field = $iIndex
	EndIf

	If ($iDataType <> Null) Then
		If Not __LO_IntIsBetween($iDataType, $LOC_SORT_DATA_TYPE_AUTO, $LOC_SORT_DATA_TYPE_ALPHANUMERIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tSortField.FieldType = $iDataType
	EndIf

	If ($bAscending <> Null) Then
		If Not IsBool($bAscending) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tSortField.IsAscending = $bAscending
	EndIf

	If ($bCaseSensitive <> Null) Then
		If Not IsBool($bCaseSensitive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tSortField.IsCaseSensitive = $bCaseSensitive
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_SortFieldModify
