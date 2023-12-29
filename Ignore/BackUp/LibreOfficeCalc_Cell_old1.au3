#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#Tidy_Parameters=/sf
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Internal.au3"

; Other includes for Calc


; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, Removing, etc. L.O. Calc document Cells.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_CellFormula
; _LOCalc_CellGetType
; _LOCalc_CellString
; _LOCalc_CellValue
; ===============================================================================================================================




; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellRangeNumbers
; Description ...: Set or Retrieve Numbers in a Range.
; Syntax ........: _LOCalc_CellRangeNumbers(ByRef $oCellRange[, $aanNumbers = Null])
; Parameters ....: $oCellRange          - [in/out] an object. A cell or cell range to set or retrieve number values for.  A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $aanNumbers          - [optional] an array of arrays containing general numbers. Default is Null. An Array of Arrays containing numbers to fill the range with. See remarks.
; Return values .: Success: 1 or Array
;				   Failure: 0 or ? and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCellRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $aanNumbers not an Array.
;				   @Error 1 @Extended 3 Return 0 = $aanNumbers array contains less or more elements than number of rows contained in the cell range.
;				   @Error 1 @Extended 4 Return ? = Element of $aanNumbers does not contain an array. Returning array element number of $aanNumbers containing error.
;				   @Error 1 @Extended 5 Return ? = Array contained in $aanNumbers has less or more elements than number of columns in the cell range. Returning array element number of $aanNumbers containing faulty array.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve array of Numerical Data contained in the Cell Range.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Start of Row from Cell Range.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve End of Row from Cell Range.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve Start of Column from Cell Range.
;				   @Error 2 @Extended 5 Return 0 = Failed to retrieve End of Column from Cell Range.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Values were successfully set for the cell range.
;				   @Error 0 @Extended 1 Return Array of Arrays = Success. $aanNumbers set to Null, returning an array containing arrays, which contain any numerical content contained in the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will return only numbers contained in the cell range when $aanNumbers is called with Null keyword. Array will be an array of arrays. The internal arrays will contain blank cells or numbers, depending on cell content.
;				   $aanNumbers must be an array containing arrays. The main Array's element count must match the row count contained in the Cell Range, and each internal Array's element count must match the column count of the Cell Range it is to fill.
;				   Any data previously contained in the Cell Range will be overwritten.
;				   All array elements must contain numbers.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellRangeNumbers(ByRef $oCellRange, $aanNumbers = Null)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $iStart, $iEnd

	If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($aanNumbers = Null) Then
		$aanNumbers = $oCellRange.getData()
		If Not IsArray($aanNumbers) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		Return SetError($__LO_STATUS_SUCCESS, 1, $aanNumbers)
		EndIf

	If Not IsArray($aanNumbers) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

; Determine if the Array is sized appropriately
	$iStart = $oCellRange.RangeAddress.StartRow()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$iEnd = $oCellRange.RangeAddress.EndRow()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If (UBound($aanNumbers) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

$iStart = $oCellRange.RangeAddress.StartColumn()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$iEnd = $oCellRange.RangeAddress.EndColumn()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)

For $i = 0 To UBound($aanNumbers) - 1
If Not IsArray($aanNumbers[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
If (UBound($aanNumbers[$i]) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
Next

$oCellRange.setData($aanNumbers)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellRangeFormula
; Description ...: Set or Retrieve Formulas in a Range.
; Syntax ........: _LOCalc_CellRangeFormula(ByRef $oCellRange[, $aasFormulas = Null])
; Parameters ....: $oCellRange          - [in/out] an object. A Cell or Cell Range to set or retrieve Formulas for. A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $aasFormulas         - [optional] an array or arrays containing strings. Default is Null. An Array of Arrays containing formula strings to fill the range with. See remarks.
; Return values .: Success: 1 or Array
;				   Failure: 0 or ? and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCellRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $aasFormulas not an Array.
;				   @Error 1 @Extended 3 Return 0 = $aasFormulas array contains less or more elements than number of rows contained in the cell range.
;				   @Error 1 @Extended 4 Return ? = Element of $aasFormulas does not contain an array. Returning array element number of $aasFormulas containing error.
;				   @Error 1 @Extended 5 Return ? = Array contained in $aasFormulas has less or more elements than number of columns in the cell range. Returning array element number of $aasFormulas containing faulty array.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve array of Formula Data contained in the Cell Range.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Start of Row from Cell Range.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve End of Row from Cell Range.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve Start of Column from Cell Range.
;				   @Error 2 @Extended 5 Return 0 = Failed to retrieve End of Column from Cell Range.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Formulas were successfully set for the cell range.
;				   @Error 0 @Extended 1 Return Array of Arrays = Success. $aasFormulas set to Null, returning an array containing arrays, which contain any Formula content contained in the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will return only formulas contained in the cell range when $aasFormulas is called with Null keyword. Array will be an array of arrays. The internal arrays will contain blank cells or formula strings, depending on cell content.
;				   $aasFormulas must be an array containing arrays. The main Array's element count must match the row count contained in the Cell Range, and each internal Array's element count must match the column count of the Cell Range it is to fill.
;				   Any data previously contained in the Cell Range will be overwritten.
;				   All array elements must contain strings, blank or otherwise.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellRangeFormula(ByRef $oCellRange, $aasFormulas = Null)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $iStart, $iEnd

	If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($aasFormulas = Null) Then
		$aasFormulas = $oCellRange.getFormulaArray()
		If Not IsArray($aasFormulas) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		Return SetError($__LO_STATUS_SUCCESS, 1, $aasFormulas)
		EndIf

	If Not IsArray($aasFormulas) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

; Determine if the Array is sized appropriately
	$iStart = $oCellRange.RangeAddress.StartRow()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$iEnd = $oCellRange.RangeAddress.EndRow()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If (UBound($aasFormulas) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

$iStart = $oCellRange.RangeAddress.StartColumn()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$iEnd = $oCellRange.RangeAddress.EndColumn()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)

For $i = 0 To UBound($aasFormulas) - 1
If Not IsArray($aasFormulas[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
If (UBound($aasFormulas[$i]) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
Next

$oCellRange.setFormulaArray($aasFormulas)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellRangeData
; Description ...: Set or Retrieve Data in a Range.
; Syntax ........: _LOCalc_CellRangeData(ByRef $oCellRange[, $aavData = Null])
; Parameters ....: $oCellRange          - [in/out] an object. The Cell or Cell Range to set or retrieve data . A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $aavData             - [optional] an array of Arrays containing variants. Default is Null. An Array of Arrays containing data, strings or numbers, to fill the range with. See remarks.
; Return values .: Success: 1 or Array
;				   Failure: 0 or ? and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCellRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $aavData not an Array.
;				   @Error 1 @Extended 3 Return 0 = $aavData array contains less or more elements than number of rows contained in the cell range.
;				   @Error 1 @Extended 4 Return ? = Element of $aavData does not contain an array. Returning array element number of $aavData containing error.
;				   @Error 1 @Extended 5 Return ? = Array contained in $aavData has less or more elements than number of columns in the cell range. Returning array element number of $aavData containing faulty array.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve array of Data contained in the Cell Range.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Start of Row from Cell Range.
;				   @Error 2 @Extended 3 Return 0 = Failed to retrieve End of Row from Cell Range.
;				   @Error 2 @Extended 4 Return 0 = Failed to retrieve Start of Column from Cell Range.
;				   @Error 2 @Extended 5 Return 0 = Failed to retrieve End of Column from Cell Range.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Data was successfully set for the cell range.
;				   @Error 0 @Extended 1 Return Array of Arrays = Success. $aavData set to Null, returning an array containing arrays, which contain any data content contained in the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will return Strings and Numbers contained in the cell range when $aavData is called with Null keyword. Array will be an array of arrays. The internal arrays will contain numerical or string data, depending on cell content.
;				   $aavData must be an array containing arrays. The main Array's element count must match the row count contained in the Cell Range, and each internal Array's element count must match the column count of the Cell Range it is to fill.
;				   Any data previously contained in the Cell Range will be overwritten.
;				   All array elements must contain appropriate data, strings or numbers.
;				   Formulas will be inserted as strings only, and will not be valid.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellRangeData(ByRef $oCellRange, $aavData = Null)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $iStart, $iEnd

	If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($aavData = Null) Then
		$aavData = $oCellRange.getDataArray()
		If Not IsArray($aavData) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		Return SetError($__LO_STATUS_SUCCESS, 1, $aavData)
		EndIf

	If Not IsArray($aavData) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

; Determine if the Array is sized appropriately
	$iStart = $oCellRange.RangeAddress.StartRow()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$iEnd = $oCellRange.RangeAddress.EndRow()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If (UBound($aavData) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

$iStart = $oCellRange.RangeAddress.StartColumn()
	If Not IsInt($iStart) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

	$iEnd = $oCellRange.RangeAddress.EndColumn()
	If Not IsInt($iEnd) Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)

For $i = 0 To UBound($aavData) - 1
If Not IsArray($aavData[$i]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)
If (UBound($aavData[$i]) <> ($iEnd - $iStart + 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
Next

$oCellRange.setDataArray($aavData)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc




; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellQueryContents
; Description ...: Query a Cell or Cell range for specific cell contents.
; Syntax ........: _LOCalc_CellQueryContents(ByRef $oCell, $iFlags)
; Parameters ....: $oCell               - [in/out] an object. A Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iFlags              - an integer value (1-1023). The Cell content type flag. Can be BitOR'd together. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFlags not an Integer, less than 1 or greater than 1023. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query cell content.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Empty cells in the range may be skipped depending on the flag used. For instance, when querying for styles, the returned ranges may not include empty cells even if styles are applied to those cells.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellQueryContents(ByRef $oCell, $iFlags)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oRanges
Local $aoRanges[0]

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iFlags, $LOC_CELL_FLAG_VALUE, $LOC_CELL_FLAG_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

$oRanges = $oCell.queryContentCells($iFlags)
If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

$aoRanges = $oRanges.getRangeAddresses()
If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

For $i = 0 To UBound($aoRanges) - 1
$aoRanges[$i] = $oCell.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellQueryDependents
; Description ...: Query a Cell or Cell Range for Dependents.
; Syntax ........: _LOCalc_CellQueryDependents(ByRef $oCell[, $bRecursive = False])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bRecursive          - [optional] a boolean value. Default is False. If True, the query is repeated for each found cell.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bRecursive not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query cell dependents.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Dependent cells are cells which reference cells in this range. If $bRecursive is True, repeats query with all found cells (finds dependents of dependents, and so on).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellQueryDependents(ByRef $oCell, $bRecursive = False)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oRanges
Local $aoRanges[0]

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bRecursive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

$oRanges = $oCell.queryDependents($bRecursive)
If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

$aoRanges = $oRanges.getRangeAddresses()
If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

For $i = 0 To UBound($aoRanges) - 1
$aoRanges[$i] = $oCell.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellQueryPrecedents
; Description ...: Query a Cell or Cell Range for Precedents.
; Syntax ........: _LOCalc_CellQueryPrecedents(ByRef $oCell[, $bRecursive = False])
; Parameters ....: $oCell               - [in/out] an object. A Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bRecursive          - [optional] a boolean value. Default is False. If True, the query is repeated for each found cell.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bRecursive not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query cell precedents.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Precedent cells are cells which are referenced by cells in this range. If $bRecursive is True, repeats query with all found cells (finds precedents of precedents, and so on).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellQueryPrecedents(ByRef $oCell, $bRecursive = False)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oRanges
Local $aoRanges[0]

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bRecursive) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

$oRanges = $oCell.queryPrecedents($bRecursive)
If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

$aoRanges = $oRanges.getRangeAddresses()
If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

For $i = 0 To UBound($aoRanges) - 1
$aoRanges[$i] = $oCell.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellQueryVisible
; Description ...: Query a Cell or Cell Range for visible cells.
; Syntax ........: _LOCalc_CellQueryVisible(ByRef $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query for visible cell.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellQueryVisible(ByRef $oCell)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oRanges
Local $aoRanges[0]

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

$oRanges = $oCell.queryVisibleCells()
If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

$aoRanges = $oRanges.getRangeAddresses()
If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

For $i = 0 To UBound($aoRanges) - 1
$aoRanges[$i] = $oCell.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellQueryEmpty
; Description ...: Query a Cell or Cell Range for empty cells.
; Syntax ........: _LOCalc_CellQueryEmpty(ByRef $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query for empty cells.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellQueryEmpty(ByRef $oCell)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oRanges
Local $aoRanges[0]

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

$oRanges = $oCell.queryEmptyCells()
If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

$aoRanges = $oRanges.getRangeAddresses()
If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

For $i = 0 To UBound($aoRanges) - 1
$aoRanges[$i] = $oCell.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellQueryFormula
; Description ...: Query a Cell or Cell Range for formulas having a specific result.
; Syntax ........: _LOCalc_CellQueryFormula(ByRef $oCell, $iResultType)
; Parameters ....: $oCell               - [in/out] an object. A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iResultType         - an integer value (1-7). The Formula result type. Can be BitOR'd together. See Constants $LOC_FORMULA_RESULT_TYPE_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iResultType not an Integer, less than 1, or greater than 7. See Constants $LOC_FORMULA_RESULT_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query cell formula results.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellQueryFormula(ByRef $oCell, $iResultType)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oRanges
Local $aoRanges[0]

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iResultType, $LOC_FORMULA_RESULT_TYPE_VALUE, $LOC_FORMULA_RESULT_TYPE_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

$oRanges = $oCell.queryFormulaCells($iResultType)
If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

$aoRanges = $oRanges.getRangeAddresses()
If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

For $i = 0 To UBound($aoRanges) - 1
$aoRanges[$i] = $oCell.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellQueryColumnDiff
; Description ...: Query a Cell Range for differences on each column based on a specific row.
; Syntax ........: _LOCalc_CellQueryColumnDiff(ByRef $oCellRange, $oCellToCompare)
; Parameters ....: $oCellRange          - [in/out] an object. A Cell Range to look for differences in. A Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oCellToCompare      - an object. A single Cell object (not a range) returned by a previous _LOCalc_SheetGetCellByName, or _LOCalc_SheetGetCellByPosition function. The Row this cell is located in will be used for the query.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCellRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCellToCompare not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCellToCompare is not a single cell, cell ranges are not supported.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Cell Address Struct from $oCellToCompare.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query column differences.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Looks for differences per column in the range, comparing the column to the value in the row $oCellToCompare is located. OOME 4.1. pg 488/489
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellQueryColumnDiff(ByRef $oCellRange, $oCellToCompare)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oRanges
Local $tCellAddr
Local $aoRanges[0]

	If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCellToCompare) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not ($oCellToCompare.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

$tCellAddr = $oCellToCompare.CellAddress()
If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

$oRanges = $oCellRange.queryColumnDifferences($tCellAddr)
If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

$aoRanges = $oRanges.getRangeAddresses()
If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

For $i = 0 To UBound($aoRanges) - 1
$aoRanges[$i] = $oCellRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellQueryRowDiff
; Description ...: Query a Cell Range for differences on each row based on a specific column.
; Syntax ........: _LOCalc_CellQueryRowDiff(ByRef $oCellRange, $oCellToCompare)
; Parameters ....: $oCellRange          - [in/out] an object. A Cell Range to look for differences in. A Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oCellToCompare      - an object. A single Cell object (not a range) returned by a previous _LOCalc_SheetGetCellByName, or _LOCalc_SheetGetCellByPosition function. The Column this cell is located in will be used for the query.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCellRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCellToCompare not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCellToCompare is not a single cell, cell ranges are not supported.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Cell Address Struct from $oCellToCompare.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query row differences.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Looks for differences per row in the range, comparing the row to the value in the column $oCellToCompare is located. OOME 4.1. pg 488/489
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellQueryRowDiff(ByRef $oCellRange, $oCellToCompare)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oRanges
Local $tCellAddr
Local $aoRanges[0]

	If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCellToCompare) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not ($oCellToCompare.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

$tCellAddr = $oCellToCompare.CellAddress()
If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

$oRanges = $oCellRange.queryRowDifferences($tCellAddr)
If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

$aoRanges = $oRanges.getRangeAddresses()
If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

For $i = 0 To UBound($aoRanges) - 1
$aoRanges[$i] = $oCellRange.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellQueryIntersection
; Description ...: Retrieve an array of cell ranges that intersect with a certain cell range.
; Syntax ........: _LOCalc_CellQueryIntersection(ByRef $oCellRange, $oCell)
; Parameters ....: $oCellRange          - [in/out] an object. A Cell range that contains the cell or cell range called in $oCell. A Cell range to delete.  A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oCell               - an object. A Cell or Cell Range located inside of the cell range called in $oCellRange. A Cell or Cell range to delete.  A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCellRange not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Range Address Struct from $oCell.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to query cell range intersections.
;				   @Error 3 @Extended 2 Return 0 = Failed to retrieve query result cell addresses.
;				   @Error 3 @Extended 3 Return 0 = Failed to retrieve cell range Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning query results array of Cell Range Objects. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOCalc_CellQueryIntersection(ByRef $oCellRange, $oCell)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oRanges
Local $tRangeAddr
Local $aoRanges[0]

	If Not IsObj($oCellRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

$tRangeAddr = $oCell.RangeAddress()
If Not IsObj($tRangeAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

$oRanges = $oCellRange.queryIntersection($tRangeAddr)
If Not IsObj($oRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

$aoRanges = $oRanges.getRangeAddresses()
If Not IsArray($aoRanges) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

For $i = 0 To UBound($aoRanges) - 1
$aoRanges[$i] = $oCell.Spreadsheet.getCellRangeByPosition($aoRanges[$i].StartColumn(), $aoRanges[$i].StartRow(), $aoRanges[$i].EndColumn(), $aoRanges[$i].EndRow())
If Not IsObj($aoRanges[$i]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($aoRanges), $aoRanges)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellRangeInsert
; Description ...: Insert blank cells at a Cell Range.
; Syntax ........: _LOCalc_CellRangeInsert(ByRef $oSheet, $oCell, $iMode)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oCell               - an object. A Cell or Cell Range to insert new blank cells at.  A Cell or Cell range to delete.  A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iMode               - an integer value (0-4). The Cell Insertion Mode. See Constants $LOC_CELL_INSERT_MODE_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iMode not an Integer, less than 0, or greater than 4. See Constants $LOC_CELL_INSERT_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Range Address Struct from $oCell.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Blank cells were successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: The new range of cells inserted will be the same size as the range called in $oCell.
;				   Non-Empty cells cannot be moved off of the sheet.
;				   This function will silently fail if the insertion will cause an array formula to be split -- OOME. 4.1., Page 509.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellRangeInsert(ByRef $oSheet, $oCell, $iMode)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tCellAddr

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iMode, $LOC_CELL_INSERT_MODE_NONE, $LOC_CELL_INSERT_MODE_COLUMNS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddr = $oCell.RangeAddress()
If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oSheet.insertCells($tCellAddr, $iMode)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellRangeDelete
; Description ...: Delete a Range of cell contents and reposition surrounding cells.
; Syntax ........: _LOCalc_CellRangeDelete(ByRef $oSheet, $oCell, $iMode)
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oCell               - an object. A Cell or Cell range to delete.  A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iMode               - an integer value (0-4). The Cell Deletion Mode. See Constants $LOC_CELL_DELETE_MODE_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iMode not an Integer, less than 0, or greater than 4. See Constants $LOC_CELL_DELETE_MODE_* as defined in LibreOfficeCalc_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Range Address Struct from $oCell.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Cell range was successfully cleared.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function will silently fail if the deletion will cause an array formula to be split -- OOME. 4.1., Page 509.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellRangeDelete(ByRef $oSheet, $oCell, $iMode)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tCellAddr

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOCalc_IntIsBetween($iMode, $LOC_CELL_DELETE_MODE_NONE, $LOC_CELL_DELETE_MODE_COLUMNS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellAddr = $oCell.RangeAddress()
If Not IsObj($tCellAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oSheet.removeRange($tCellAddr, $iMode)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellClearContents
; Description ...: Clear specific cell contents in a range.
; Syntax ........: _LOCalc_CellClearContents(ByRef $oCell, $iFlags)
; Parameters ....: $oCell               - [in/out] an object. A Cell or Cell range to clear the contents of.  A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iFlags              - an integer value (1-1023). The Cell Content type to delete. Can be BitOR'd together. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFlags not an Integer, less than 1, or greater than 1023. See Constants $LOC_CELL_FLAG_* as defined in LibreOfficeCalc_Constants.au3.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Contents specified was successfully cleared from the cell range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellClearContents(ByRef $oCell, $iFlags)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOCalc_IntIsBetween($iFlags, $LOC_CELL_FLAG_VALUE, $LOC_CELL_FLAG_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

$oCell.clearContents($iFlags)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellRangeCopyMove
; Description ...: Copy or Move a Cell or Cell Range to another range.
; Syntax ........: _LOCalc_CellRangeCopyMove(ByRef $oSheet, ByRef $oCellSrc, ByRef $oCellDest[, $bMove = False])
; Parameters ....: $oSheet              - [in/out] an object. A Sheet object returned by a previous _LOCalc_DocSheetAdd, _LOCalc_SheetGetActive, _LOCalc_SheetCopy, or _LOCalc_SheetGetObjByName function.
;                  $oCellSrc            - [in/out] an object. The Cell or Cell Range to copy or move from. A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $oCellDest           - [in/out] an object. The Cell or Cell Range to copy or move to. A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bMove               - [optional] a boolean value. Default is False. If True, the cell range is moved to the destination. If False, the Cell Range is only copied.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSheet not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCellSrc not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCellDest not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bMove not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Source Cell Range Address.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Destination Cell Range Address.
;				   @Error 2 @Extended 3 Return 0 = Failed to create "com.sun.star.table.CellAddress" Struct.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Cell or Cell range was successfully copied to destination.
;				   @Error 0 @Extended 1 Return 1 = Success. Cell or Cell range was successfully moved to destination.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Destination Range can be on different Sheet from Source.
;				   $oSheet is the source sheet where $oCellSrc is located.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellRangeCopyMove(ByRef $oSheet, ByRef $oCellSrc, ByRef $oCellDest, $bMove = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tCellSrcAddr, $tInputCellDestAddr, $tCellDestAddr

	If Not IsObj($oSheet) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCellSrc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oCellDest) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bMove) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$tCellSrcAddr = $oCellSrc.RangeAddress()
If Not IsObj($tCellSrcAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tInputCellDestAddr = $oCellDest.RangeAddress()
If Not IsObj($tInputCellDestAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

$tCellDestAddr = __LOCalc_CreateStruct("com.sun.star.table.CellAddress")
If Not IsObj($tCellDestAddr) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

$tCellDestAddr.Sheet = $tInputCellDestAddr.Sheet()
$tCellDestAddr.Column = $tInputCellDestAddr.StartColumn()
$tCellDestAddr.Row = $tInputCellDestAddr.StartRow()

If $bMove Then
$oSheet.MoveRange($tCellDestAddr, $tCellSrcAddr)
Return SetError($__LO_STATUS_SUCCESS, 1, 1)

Else
	$oSheet.CopyRange($tCellDestAddr, $tCellSrcAddr)
	Return SetError($__LO_STATUS_SUCCESS, 0, 1)

EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellBackColor
; Description ...: Set or Retrieve the Cell or Cell Range Background color.
; Syntax ........: _LOCalc_CellBackColor(ByRef $oCell[, $iBackColor = Null[, $bBackTransparent = Null]])
; Parameters ....: $oCell               - [in/out] an object. A Cell or Cell Range object returned by a previous _LOCalc_SheetGetCellByName, _LOCalc_SheetGetCellByPosition, _LOCalc_SheetColumnGetObjByPosition, _LOCalc_SheetColumnGetObjByName, _LOcalc_SheetRowGetObj, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The Cell background color as a Long Integer. Set to $LOC_COLOR_OFF(-1) to disable Background color. Can also be one of the constants $LOC_COLOR_* as defined in LibreOfficeCalc_Constants.au3
; Return values .: Success: 1 or Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iBackColor not an Integer, set less than -1 or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iBackColor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current Cell or Cell Range background color setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current setting.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOCalc_ConvertColorToLong, _LOCalc_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellBackColor(ByRef $oCell, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOCalc_VarsAreNull($iBackColor) Then Return SetError($__LO_STATUS_SUCCESS, 0, $oCell.CellBackColor())

		If Not __LOCalc_IntIsBetween($iBackColor, $LOC_COLOR_OFF, $LOC_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oCell.CellBackColor = $iBackColor
		$iError = ($oCell.CellBackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1)) ; Error setting color.

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc




; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellFormula
; Description ...: Set or Retrieve a Cell's Formula.
; Syntax ........: _LOCalc_CellFormula(ByRef $oCell[, $sFormula = Null])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by previous _LOCalc_SheetGetCellByName, or _LOCalc_SheetGetCellByPosition function.
;                  $sFormula            - [optional] a string value. Default is Null. The Formula to set the Cell to. Overwrites any previous data.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;				   @Error 1 @Extended 3 Return 0 = $sFormula not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $sFormula
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning the Cell's current formula as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: Only individual cells are supported, not cell ranges.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current Cell content.
; Related .......: _LOCalc_CellGetType, _LOCalc_CellString, _LOCalc_CellValue
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellFormula(ByRef $oCell, $sFormula = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	If ($sFormula = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oCell.getFormula())

	If Not IsString($sFormula) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCell.setFormula($sFormula)
	If ($oCell.getFormula() <> $sFormula) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CellFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellGetType
; Description ...: Retrieve the type of data that is contained in a Cell.
; Syntax ........: _LOCalc_CellGetType(ByRef $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by previous _LOCalc_SheetGetCellByName, or _LOCalc_SheetGetCellByPosition function.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Data Type.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Returning type of data contained in the Cell. Return value will be one of the constants, $LOC_CELL_TYPE_* as defined in LibreOfficeCalc_Constants.au3
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: Only individual cells are supported, not cell ranges.
; Related .......: _LOCalc_CellString, _LOCalc_CellFormula, _LOCalc_CellValue
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellGetType(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCellType

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	$iCellType = $oCell.getType()
	If Not IsInt($iCellType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCellType)
EndFunc   ;==>_LOCalc_CellGetType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellString
; Description ...: Set or Retrieve a Cell's Text content.
; Syntax ........: _LOCalc_CellString(ByRef $oCell[, $sText = Null])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by previous _LOCalc_SheetGetCellByName, or _LOCalc_SheetGetCellByPosition function.
;                  $sText               - [optional] a string value. Default is Null. The Text to set the Cell contents to. Overwrites any previous data.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;				   @Error 1 @Extended 3 Return 0 = $sText not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $sText
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning the Cell's contents as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: Only individual cells are supported, not cell ranges.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current Cell content.
; Related .......: _LOCalc_CellGetType, _LOCalc_CellFormula, _LOCalc_CellValue
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellString(ByRef $oCell, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	If ($sText = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oCell.getString())

	If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCell.setString($sText)
	If ($oCell.getString() <> $sText) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CellString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_CellValue
; Description ...: Set or Retrieve a Cell's Value.
; Syntax ........: _LOCalc_CellValue(ByRef $oCell[, $nValue = Null])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by previous _LOCalc_SheetGetCellByName, or _LOCalc_SheetGetCellByPosition function.
;                  $nValue              - [optional] a general number value. Default is Null. The Value to set the Cell to. Overwrites any previous data.
; Return values .: Success: 1 or Number.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCell is a Cell Range and is not supported.
;				   @Error 1 @Extended 3 Return 0 = $nValue not a Number.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;				   |								1 = Error setting $nValue
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Number = Success. All optional parameters were set to Null, returning the Cell's current number value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: Only individual cells are supported, not cell ranges.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current Cell content.
; Related .......: _LOCalc_CellGetType, _LOCalc_CellString, _LOCalc_CellFormula
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_CellValue(ByRef $oCell, $nValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oCell.supportsService("com.sun.star.sheet.SheetCell")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Only single cells supported.

	If ($nValue = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oCell.getValue())

	If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCell.setValue($nValue)
	If ($oCell.getValue() <> $nValue) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_CellValue
