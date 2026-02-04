#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Base
#include "LibreOfficeBase_Constants.au3"
#include "LibreOfficeBase_Internal.au3"

; Other includes for Base

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Functions used for creating, modifying and executing SQL Statements in LibreOffice Base.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_SQLResultColumnMetaDataQuery
; _LOBase_SQLResultColumnsGetCount
; _LOBase_SQLResultColumnsGetNames
; _LOBase_SQLResultCursorMove
; _LOBase_SQLResultCursorQuery
; _LOBase_SQLResultRowModify
; _LOBase_SQLResultRowQuery
; _LOBase_SQLResultRowRead
; _LOBase_SQLResultRowRefresh
; _LOBase_SQLResultRowUpdate
; _LOBase_SQLStatementCreate
; _LOBase_SQLStatementExecuteQuery
; _LOBase_SQLStatementExecuteUpdate
; _LOBase_SQLStatementPreparedSetData
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLResultColumnMetaDataQuery
; Description ...: Query a Result set for current column status or settings.
; Syntax ........: _LOBase_SQLResultColumnMetaDataQuery(ByRef $oResult, $iColumn, $iQuery)
; Parameters ....: $oResult             - [in/out] an object. A Result Set object returned by a previous _LOBase_SQLStatementExecuteQuery, _LOBase_QueryUIGetRowSet, or _LOBase_TableUIGetRowSet function.
;                  $iColumn             - an integer value. The column to perform the Query on. 1 based.
;                  $iQuery              - an integer value (0-18). The Query command to perform. See Constants, $LOB_RESULT_METADATA_QUERY_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: Variable
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oResult not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oResult not a Result Set Object.
;                  @Error 1 @Extended 3 Return 0 = $iColumn not an Integer, less than 1 or greater than number of Columns contained in Result Set.
;                  @Error 1 @Extended 4 Return 0 = $iQuery not an Integer, less than 0 or greater than 18. See Constants, $LOB_RESULT_METADATA_QUERY_* as defined in LibreOfficeBase_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Column Count.
;                  @Error 3 @Extended 2 Return 0 = Failed to Execute Query.
;                  --Success--
;                  @Error 0 @Extended 0 Return Variable = Success. Returning Query result. See Query description for expected return type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLResultColumnMetaDataQuery(ByRef $oResult, $iColumn, $iQuery)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $asCommands[19]
	Local $iCount

	$asCommands[$LOB_RESULT_METADATA_QUERY_GET_CATALOG_NAME] = ".getCatalogName"
	$asCommands[$LOB_RESULT_METADATA_QUERY_GET_DECIMAL_PLACE] = ".getScale"
	$asCommands[$LOB_RESULT_METADATA_QUERY_GET_DISP_SIZE] = ".getColumnDisplaySize"
	$asCommands[$LOB_RESULT_METADATA_QUERY_GET_LABEL] = ".getColumnLabel"
	$asCommands[$LOB_RESULT_METADATA_QUERY_GET_LENGTH] = ".getPrecision"
	$asCommands[$LOB_RESULT_METADATA_QUERY_GET_NAME] = ".getColumnName"
	$asCommands[$LOB_RESULT_METADATA_QUERY_GET_SCHEMA_NAME] = ".getSchemaName"
	$asCommands[$LOB_RESULT_METADATA_QUERY_GET_TABLE_NAME] = ".getTableName"
	$asCommands[$LOB_RESULT_METADATA_QUERY_GET_TYPE] = ".getColumnType"
	$asCommands[$LOB_RESULT_METADATA_QUERY_GET_TYPE_NAME] = ".getColumnTypeName"
	$asCommands[$LOB_RESULT_METADATA_QUERY_IS_AUTO_VALUE] = ".isAutoIncrement"
	$asCommands[$LOB_RESULT_METADATA_QUERY_IS_CASE_SENSITIVE] = ".isCaseSensitive"
	$asCommands[$LOB_RESULT_METADATA_QUERY_IS_CURRENCY] = ".isCurrency"
	$asCommands[$LOB_RESULT_METADATA_QUERY_IS_NULLABLE] = ".isNullable"
	$asCommands[$LOB_RESULT_METADATA_QUERY_IS_READ_ONLY] = ".isReadOnly"
	$asCommands[$LOB_RESULT_METADATA_QUERY_IS_SEARCHABLE] = ".isSearchable"
	$asCommands[$LOB_RESULT_METADATA_QUERY_IS_SIGNED] = ".isSigned"
	$asCommands[$LOB_RESULT_METADATA_QUERY_IS_WRITABLE] = ".isWritable"
	$asCommands[$LOB_RESULT_METADATA_QUERY_IS_WRITABLE_DEFINITE] = ".isDefinitelyWritable"

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oResult.supportsService("com.sun.star.sdb.ResultSet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iCount = $oResult.Columns.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iColumn, 1, $iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iQuery, $LOB_RESULT_METADATA_QUERY_GET_CATALOG_NAME, $LOB_RESULT_METADATA_QUERY_IS_WRITABLE_DEFINITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	Switch $iQuery
		Case $LOB_RESULT_METADATA_QUERY_GET_CATALOG_NAME, $LOB_RESULT_METADATA_QUERY_GET_SCHEMA_NAME, $LOB_RESULT_METADATA_QUERY_GET_TABLE_NAME, $LOB_RESULT_METADATA_QUERY_GET_LABEL, _
				$LOB_RESULT_METADATA_QUERY_GET_NAME, $LOB_RESULT_METADATA_QUERY_GET_TYPE_NAME
			$vReturn = Execute("$oResult.MetaData" & $asCommands[$iQuery] & "(" & $iColumn & ")")
			If @error Or Not IsString($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Case $LOB_RESULT_METADATA_QUERY_GET_DISP_SIZE, $LOB_RESULT_METADATA_QUERY_GET_TYPE, $LOB_RESULT_METADATA_QUERY_GET_LENGTH, $LOB_RESULT_METADATA_QUERY_GET_DECIMAL_PLACE, _
				$LOB_RESULT_METADATA_QUERY_IS_NULLABLE
			$vReturn = Execute("$oResult.MetaData" & $asCommands[$iQuery] & "(" & $iColumn & ")")
			If @error Or Not IsInt($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Case $LOB_RESULT_METADATA_QUERY_IS_AUTO_VALUE, $LOB_RESULT_METADATA_QUERY_IS_CASE_SENSITIVE, $LOB_RESULT_METADATA_QUERY_IS_CURRENCY, $LOB_RESULT_METADATA_QUERY_IS_READ_ONLY, _
				$LOB_RESULT_METADATA_QUERY_IS_SEARCHABLE, $LOB_RESULT_METADATA_QUERY_IS_SIGNED, $LOB_RESULT_METADATA_QUERY_IS_WRITABLE, $LOB_RESULT_METADATA_QUERY_IS_WRITABLE_DEFINITE
			$vReturn = Execute("$oResult.MetaData" & $asCommands[$iQuery] & "(" & $iColumn & ")")
			If @error Or Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, $vReturn)
EndFunc   ;==>_LOBase_SQLResultColumnMetaDataQuery

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLResultColumnsGetCount
; Description ...: Retrieve a count of Columns returned in a Result Set.
; Syntax ........: _LOBase_SQLResultColumnsGetCount(ByRef $oResult)
; Parameters ....: $oResult             - [in/out] an object. A Result Set object returned by a previous _LOBase_SQLStatementExecuteQuery, _LOBase_QueryUIGetRowSet, or _LOBase_TableUIGetRowSet function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oResult not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oResult not a Result Set Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Column Count.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Columns contained in the Result Set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLResultColumnsGetCount(ByRef $oResult)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iReturn

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oResult.supportsService("com.sun.star.sdb.ResultSet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iReturn = $oResult.MetaData.ColumnCount()
	If Not IsInt($iReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iReturn)
EndFunc   ;==>_LOBase_SQLResultColumnsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLResultColumnsGetNames
; Description ...: Retrieve an Array of Column names Returned in a Result Set.
; Syntax ........: _LOBase_SQLResultColumnsGetNames(ByRef $oResult)
; Parameters ....: $oResult             - [in/out] an object. A Result Set object returned by a previous _LOBase_SQLStatementExecuteQuery, _LOBase_QueryUIGetRowSet, or _LOBase_TableUIGetRowSet function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oResult not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oResult not a Result Set Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Array of Column Names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Column Names contained in the Result Set. @Extended is set to the number of Elements contained in the Array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLResultColumnsGetNames(ByRef $oResult)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asReturn[0]

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oResult.supportsService("com.sun.star.sdb.ResultSet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$asReturn = $oResult.Columns.ElementNames()
	If Not IsArray($asReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asReturn), $asReturn)
EndFunc   ;==>_LOBase_SQLResultColumnsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLResultCursorMove
; Description ...: Move the Result Set Cursor within the Result Set.
; Syntax ........: _LOBase_SQLResultCursorMove(ByRef $oResult, $iMove[, $iNumber = Null])
; Parameters ....: $oResult             - [in/out] an object. A Result Set object returned by a previous _LOBase_SQLStatementExecuteQuery, _LOBase_QueryUIGetRowSet, or _LOBase_TableUIGetRowSet function.
;                  $iMove               - an integer value (0-7). The move command for the cursor. See Constants, $LOB_RESULT_CURSOR_MOVE_* as defined in LibreOfficeBase_Constants.au3.
;                  $iNumber             - [optional] an integer value. Default is Null. The Absolute row number or number of moves to go forward or backward. See Remarks.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oResult not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oResult not a Result Set object.
;                  @Error 1 @Extended 3 Return 0 = $iMove not an Integer, less than 0 or greater than 7. See Constants, $LOB_RESULT_CURSOR_MOVE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iNumber not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to process Cursor move.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning Boolean whether the move was successful (True) or not (False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iNumber is only used when calling $LOB_RESULT_CURSOR_MOVE_ABSOLUTE or $LOB_RESULT_CURSOR_MOVE_RELATIVE commands.
;                  When $LOB_RESULT_CURSOR_MOVE_RELATIVE is called, both positive and negative numbers may be used. Forwards is a positive value, and backwards is a negative value.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLResultCursorMove(ByRef $oResult, $iMove, $iNumber = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn = True
	Local $asMoves[8]
	$asMoves[$LOB_RESULT_CURSOR_MOVE_BEFORE_FIRST] = "beforeFirst"
	$asMoves[$LOB_RESULT_CURSOR_MOVE_FIRST] = "first"
	$asMoves[$LOB_RESULT_CURSOR_MOVE_PREVIOUS] = "previous"
	$asMoves[$LOB_RESULT_CURSOR_MOVE_NEXT] = "next"
	$asMoves[$LOB_RESULT_CURSOR_MOVE_LAST] = "last"
	$asMoves[$LOB_RESULT_CURSOR_MOVE_AFTER_LAST] = "afterLast"
	$asMoves[$LOB_RESULT_CURSOR_MOVE_ABSOLUTE] = "absolute"
	$asMoves[$LOB_RESULT_CURSOR_MOVE_RELATIVE] = "relative"

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oResult.supportsService("com.sun.star.sdb.ResultSet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iMove, $LOB_RESULT_CURSOR_MOVE_BEFORE_FIRST, $LOB_RESULT_CURSOR_MOVE_RELATIVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	Switch $iMove
		Case $LOB_RESULT_CURSOR_MOVE_BEFORE_FIRST, $LOB_RESULT_CURSOR_MOVE_AFTER_LAST
			Execute("$oResult." & $asMoves[$iMove] & "()")

		Case $LOB_RESULT_CURSOR_MOVE_FIRST, $LOB_RESULT_CURSOR_MOVE_PREVIOUS, $LOB_RESULT_CURSOR_MOVE_NEXT, $LOB_RESULT_CURSOR_MOVE_LAST
			$bReturn = Execute("$oResult." & $asMoves[$iMove] & "()")

		Case $LOB_RESULT_CURSOR_MOVE_ABSOLUTE, $LOB_RESULT_CURSOR_MOVE_RELATIVE
			If Not IsInt($iNumber) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$bReturn = Execute("$oResult." & $asMoves[$iMove] & "(" & $iNumber & ")")
	EndSwitch

	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_SQLResultCursorMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLResultCursorQuery
; Description ...: Perform a Query on the Result Set Cursor position.
; Syntax ........: _LOBase_SQLResultCursorQuery(ByRef $oResult, $iQuery)
; Parameters ....: $oResult             - [in/out] an object. A Result Set object returned by a previous _LOBase_SQLStatementExecuteQuery, _LOBase_QueryUIGetRowSet, or _LOBase_TableUIGetRowSet function.
;                  $iQuery              - an integer value (0-4). The Query to perform on the cursor. See Constants, $LOB_RESULT_CURSOR_QUERY_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: Boolean or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oResult not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oResult not a Result Set object.
;                  @Error 1 @Extended 3 Return 0 = $iQuery not an Integer, less than 0 or greater than 4. See Constants, $LOB_RESULT_CURSOR_QUERY_* as defined in LibreOfficeBase_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to process Cursor Query.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning cursor query result.
;                  @Error 0 @Extended 0 Return Integer = Success. Returning current row number containing the cursor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLResultCursorQuery(ByRef $oResult, $iQuery)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn
	Local $iReturn
	Local $asQuery[5]
	$asQuery[$LOB_RESULT_CURSOR_QUERY_IS_BEFORE_FIRST] = "isBeforeFirst"
	$asQuery[$LOB_RESULT_CURSOR_QUERY_IS_FIRST] = "isFirst"
	$asQuery[$LOB_RESULT_CURSOR_QUERY_IS_LAST] = "isLast"
	$asQuery[$LOB_RESULT_CURSOR_QUERY_IS_AFTER_LAST] = "isAfterLast"
	$asQuery[$LOB_RESULT_CURSOR_QUERY_GET_ROW] = "getRow"

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oResult.supportsService("com.sun.star.sdb.ResultSet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iQuery, $LOB_RESULT_CURSOR_QUERY_IS_BEFORE_FIRST, $LOB_RESULT_CURSOR_QUERY_GET_ROW) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	Switch $iQuery
		Case $LOB_RESULT_CURSOR_QUERY_IS_BEFORE_FIRST, $LOB_RESULT_CURSOR_QUERY_IS_FIRST, $LOB_RESULT_CURSOR_QUERY_IS_LAST, $LOB_RESULT_CURSOR_QUERY_IS_AFTER_LAST
			$bReturn = Execute("$oResult." & $asQuery[$iQuery] & "()")
			If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)

		Case $LOB_RESULT_CURSOR_QUERY_GET_ROW
			$iReturn = Execute("$oResult." & $asQuery[$iQuery] & "()")
			If Not IsInt($iReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			Return SetError($__LO_STATUS_SUCCESS, 1, $iReturn)
	EndSwitch
EndFunc   ;==>_LOBase_SQLResultCursorQuery

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLResultRowModify
; Description ...: Modify the values of the current Result Set Row.
; Syntax ........: _LOBase_SQLResultRowModify(ByRef $oResult, $iModify, $iColumn, $vValue)
; Parameters ....: $oResult             - [in/out] an object. A Result Set object returned by a previous _LOBase_SQLStatementExecuteQuery, _LOBase_QueryUIGetRowSet, or _LOBase_TableUIGetRowSet function.
;                  $iModify             - an integer value (0-12). The modify command for the Result Set Row. See Constants, $LOB_RESULT_ROW_MOD_* as defined in LibreOfficeBase_Constants.au3.
;                  $iColumn             - an integer value. The column to perform the Modification upon. 1 based.
;                  $vValue              - a variant value. The Value to change the column to.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oResult not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oResult not a Result Set object.
;                  @Error 1 @Extended 3 Return 0 = $iModify not an Integer, less than 0 or greater than 12. See Constants, $LOB_RESULT_ROW_MOD_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iColumn not an Integer, or less than 1.
;                  @Error 1 @Extended 5 Return 0 = $iModify called with $LOB_RESULT_ROW_MOD_BOOL and $vValue not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iModify called with $LOB_RESULT_ROW_MOD_SHORT, $LOB_RESULT_ROW_MOD_INT, or $LOB_RESULT_ROW_MOD_LONG and $vValue not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iModify called with $LOB_RESULT_ROW_MOD_FLOAT, or $LOB_RESULT_ROW_MOD_DOUBLE and $vValue not a Number.
;                  @Error 1 @Extended 8 Return 0 = $iModify called with $LOB_RESULT_ROW_MOD_STRING and $vValue not a String.
;                  @Error 1 @Extended 9 Return 0 = $iModify called with $LOB_RESULT_ROW_MOD_DATE, $LOB_RESULT_ROW_MOD_TIME, or $LOB_RESULT_ROW_MOD_TIMESTAMP and $vValue not an Object.
;                  @Error 1 @Extended 10 Return 0 = $iModify called with $LOB_RESULT_ROW_MOD_BYTE, or $LOB_RESULT_ROW_MOD_BYTES and $vValue not a Binary value.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.Date" Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.util.Time" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Modification command.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully performed the Result Row Modification command.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $vValue is ignored when calling the $LOB_RESULT_ROW_MOD_NULL command.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLResultRowModify(ByRef $oResult, $iModify, $iColumn, $vValue)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tDateTime

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oResult.supportsService("com.sun.star.sdb.ResultSet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iModify, $LOB_RESULT_ROW_MOD_NULL, $LOB_RESULT_ROW_MOD_TIMESTAMP) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iColumn, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	Switch $iModify
		Case $LOB_RESULT_ROW_MOD_NULL
			$oResult.updateNull($iColumn)

		Case $LOB_RESULT_ROW_MOD_BOOL
			If Not IsBool($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oResult.updateBoolean($iColumn, $vValue)

		Case $LOB_RESULT_ROW_MOD_SHORT
			If Not IsInt($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oResult.updateShort($iColumn, $vValue)

		Case $LOB_RESULT_ROW_MOD_INT
			If Not IsInt($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oResult.updateInt($iColumn, $vValue)

		Case $LOB_RESULT_ROW_MOD_LONG
			If Not IsInt($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oResult.updateLong($iColumn, $vValue)

		Case $LOB_RESULT_ROW_MOD_FLOAT
			If Not IsNumber($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oResult.updateFloat($iColumn, $vValue)

		Case $LOB_RESULT_ROW_MOD_DOUBLE
			If Not IsNumber($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oResult.updateDouble($iColumn, $vValue)

		Case $LOB_RESULT_ROW_MOD_STRING
			If Not IsString($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oResult.updateString($iColumn, $vValue)

		Case $LOB_RESULT_ROW_MOD_DATE
			If Not IsObj($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$tDateTime = __LO_CreateStruct("com.sun.star.util.Date")
			If Not IsObj($tDateTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tDateTime.Year = $vValue.Year()
			$tDateTime.Month = $vValue.Month()
			$tDateTime.Day = $vValue.Day()

			$oResult.updateDate($iColumn, $tDateTime)

		Case $LOB_RESULT_ROW_MOD_TIME
			If Not IsObj($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$tDateTime = __LO_CreateStruct("com.sun.star.util.Time")
			If Not IsObj($tDateTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$tDateTime.Hours = $vValue.Hours()
			$tDateTime.Minutes = $vValue.Minutes()
			$tDateTime.Seconds = $vValue.Seconds()
			$tDateTime.NanoSeconds = $vValue.NanoSeconds()
			If __LO_VersionCheck(4.1) Then $tDateTime.IsUTC = $vValue.IsUTC()

			$oResult.updateTime($iColumn, $tDateTime)

		Case $LOB_RESULT_ROW_MOD_TIMESTAMP
			If Not IsObj($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$oResult.updateTimestamp($iColumn, $vValue)

		Case $LOB_RESULT_ROW_MOD_BYTE
			If Not IsBinary($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			$oResult.updateByte($iColumn, $vValue)

		Case $LOB_RESULT_ROW_MOD_BYTES
			If Not IsBinary($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			$oResult.updateBytes($iColumn, $vValue)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_SQLResultRowModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLResultRowQuery
; Description ...: Query the status of the current Result Set Row.
; Syntax ........: _LOBase_SQLResultRowQuery(ByRef $oResult, $iQuery)
; Parameters ....: $oResult             - [in/out] an object. A Result Set object returned by a previous _LOBase_SQLStatementExecuteQuery, _LOBase_QueryUIGetRowSet, or _LOBase_TableUIGetRowSet function.
;                  $iQuery              - an integer value (0-2). The Query to perform for the current row of the Result Set. See Constants, $LOB_RESULT_ROW_QUERY_IS_ROW_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oResult not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oResult not a Result Set object.
;                  @Error 1 @Extended 3 Return 0 = $iQuery not an Integer, less than 0 or greater than 2. See Constants, $LOB_RESULT_ROW_QUERY_IS_ROW_* as defined in LibreOfficeBase_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to process query.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning result of query as a Boolean value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLResultRowQuery(ByRef $oResult, $iQuery)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn
	Local $asQuery[3]
	$asQuery[$LOB_RESULT_ROW_QUERY_IS_ROW_INSERTED] = "rowInserted"
	$asQuery[$LOB_RESULT_ROW_QUERY_IS_ROW_UPDATED] = "rowUpdated"
	$asQuery[$LOB_RESULT_ROW_QUERY_IS_ROW_DELETED] = "rowDeleted"

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oResult.supportsService("com.sun.star.sdb.ResultSet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iQuery, $LOB_RESULT_ROW_QUERY_IS_ROW_INSERTED, $LOB_RESULT_ROW_QUERY_IS_ROW_DELETED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$bReturn = Execute("$oResult." & $asQuery[$iQuery] & "()")
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_SQLResultRowQuery

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLResultRowRead
; Description ...: Read a column from the current Result Set Row.
; Syntax ........: _LOBase_SQLResultRowRead(ByRef $oResult, $iRead, $iColumn)
; Parameters ....: $oResult             - [in/out] an object. A Result Set object returned by a previous _LOBase_SQLStatementExecuteQuery, _LOBase_QueryUIGetRowSet, or _LOBase_TableUIGetRowSet function.
;                  $iRead               - an integer value (0-12). The read command to perform for the Result Set Row. See Constants, $LOB_RESULT_ROW_READ_* as defined in LibreOfficeBase_Constants.au3.
;                  $iColumn             - an integer value. The column to perform the Query for. 1 based.
; Return values .: Success: Variable
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oResult not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oResult not a Result Set object.
;                  @Error 1 @Extended 3 Return 0 = $iRead not an Integer, less than 0 or greater than 12. See Constants, $LOB_RESULT_ROW_READ_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iColumn not an Integer, or less than 1.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.DateTime" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve a Date Struct from Row Read.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve a Time Struct from Row Read.
;                  --Success--
;                  @Error 0 @Extended 0 Return Variable = Success. Successfully performed Row read, returning corresponding data type as the read command performed.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You can read most, or all values using $LOB_RESULT_ROW_READ_STRING if the returned Data type does not matter.
;                  Before you can call $LOB_RESULT_ROW_READ_WAS_NULL, a previous query has to have been already performed.
;                  When querying a Date or Time value, a Date Structure is returned, which you can then use the function _LOBase_DateStructModify to retrieve the values for.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLResultRowRead(ByRef $oResult, $iRead, $iColumn)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tDateTime
	Local $vReturn
	Local $asRead[13]
	$asRead[$LOB_RESULT_ROW_READ_STRING] = "getString"
	$asRead[$LOB_RESULT_ROW_READ_BOOL] = "getBoolean"
	$asRead[$LOB_RESULT_ROW_READ_BYTE] = "getByte"
	$asRead[$LOB_RESULT_ROW_READ_SHORT] = "getShort"
	$asRead[$LOB_RESULT_ROW_READ_INT] = "getInt"
	$asRead[$LOB_RESULT_ROW_READ_LONG] = "getLong"
	$asRead[$LOB_RESULT_ROW_READ_FLOAT] = "getFloat"
	$asRead[$LOB_RESULT_ROW_READ_DOUBLE] = "getDouble"
	$asRead[$LOB_RESULT_ROW_READ_BYTES] = "getBytes"
	$asRead[$LOB_RESULT_ROW_READ_DATE] = "getDate"
	$asRead[$LOB_RESULT_ROW_READ_TIME] = "getTime"
	$asRead[$LOB_RESULT_ROW_READ_TIMESTAMP] = "getTimestamp"
	$asRead[$LOB_RESULT_ROW_READ_WAS_NULL] = "wasNull"

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oResult.supportsService("com.sun.star.sdb.ResultSet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iRead, $LOB_RESULT_ROW_READ_STRING, $LOB_RESULT_ROW_READ_WAS_NULL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iColumn, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	Switch $iRead
		Case $LOB_RESULT_ROW_READ_WAS_NULL
			$vReturn = Execute("$oResult." & $asRead[$iRead] & "()")

		Case $LOB_RESULT_ROW_READ_DATE
			$vReturn = Execute("$oResult." & $asRead[$iRead] & "(" & $iColumn & ")")
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$tDateTime = __LO_CreateStruct("com.sun.star.util.DateTime")
			If Not IsObj($tDateTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tDateTime.Year = $vReturn.Year()
			$tDateTime.Month = $vReturn.Month()
			$tDateTime.Day = $vReturn.Day()

			$vReturn = $tDateTime

		Case $LOB_RESULT_ROW_READ_TIME
			$vReturn = Execute("$oResult." & $asRead[$iRead] & "(" & $iColumn & ")")
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tDateTime = __LO_CreateStruct("com.sun.star.util.DateTime")
			If Not IsObj($tDateTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tDateTime.Hours = $vReturn.Hours()
			$tDateTime.Minutes = $vReturn.Minutes()
			$tDateTime.Seconds = $vReturn.Seconds()
			$tDateTime.NanoSeconds = $vReturn.NanoSeconds()
			If __LO_VersionCheck(4.1) Then $tDateTime.IsUTC = $vReturn.IsUTC()

			$vReturn = $tDateTime

		Case Else
			$vReturn = Execute("$oResult." & $asRead[$iRead] & "(" & $iColumn & ")")
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, $vReturn)
EndFunc   ;==>_LOBase_SQLResultRowRead

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLResultRowRefresh
; Description ...: Read the original values back into the Result Row Set.
; Syntax ........: _LOBase_SQLResultRowRefresh(ByRef $oResult)
; Parameters ....: $oResult             - [in/out] an object. A Result Set object returned by a previous _LOBase_SQLStatementExecuteQuery, _LOBase_QueryUIGetRowSet, or _LOBase_TableUIGetRowSet function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oResult not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oResult not a Result Set object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully refreshed the Result Set Row.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLResultRowRefresh(ByRef $oResult)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oResult.supportsService("com.sun.star.sdb.ResultSet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oResult.refreshRow()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_SQLResultRowRefresh

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLResultRowUpdate
; Description ...: Perform an Update for the current Result Set Row.
; Syntax ........: _LOBase_SQLResultRowUpdate(ByRef $oResult, $iUpdate)
; Parameters ....: $oResult             - [in/out] an object. A Result Set object returned by a previous _LOBase_SQLStatementExecuteQuery, _LOBase_QueryUIGetRowSet, or _LOBase_TableUIGetRowSet function.
;                  $iUpdate             - an integer value (0-5). The Update command to perform for the current row of the Result Set. See Constants, $LOB_RESULT_ROW_UPDATE_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oResult not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oResult not a Result Set object.
;                  @Error 1 @Extended 3 Return 0 = $iUpdate not an Integer, less than 0 or greater than 5. See Constants, $LOB_RESULT_ROW_UPDATE_* as defined in LibreOfficeBase_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to process Update.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully executed Update command.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLResultRowUpdate(ByRef $oResult, $iUpdate)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asUpdate[6]
	$asUpdate[$LOB_RESULT_ROW_UPDATE_INSERT] = "insertRow"
	$asUpdate[$LOB_RESULT_ROW_UPDATE_UPDATE] = "updateRow"
	$asUpdate[$LOB_RESULT_ROW_UPDATE_DELETE] = "deleteRow"
	$asUpdate[$LOB_RESULT_ROW_UPDATE_CANCEL_UPDATE] = "cancelRowUpdates"
	$asUpdate[$LOB_RESULT_ROW_UPDATE_MOVE_TO_INSERT] = "moveToInsertRow"
	$asUpdate[$LOB_RESULT_ROW_UPDATE_MOVE_TO_CURRENT] = "moveToCurrentRow"

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oResult.supportsService("com.sun.star.sdb.ResultSet") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iUpdate, $LOB_RESULT_ROW_UPDATE_INSERT, $LOB_RESULT_ROW_UPDATE_MOVE_TO_CURRENT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	Execute("$oResult." & $asUpdate[$iUpdate] & "()")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_SQLResultRowUpdate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLStatementCreate
; Description ...: Create a Prepared SQL Statement or a SQL statement to perform an Update or a Query with.
; Syntax ........: _LOBase_SQLStatementCreate(ByRef $oConnection[, $sSQL = Null])
; Parameters ....: $oConnection         - [in/out] an object. A Statement object returned by a previous _LOBase_SQLStatementCreate function.
;                  $sSQL                - [optional] a string value. Default is Null. The SQL string to create the Prepared statement with.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $sSQL not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a prepared Statement.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a Statement.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the created Prepared Statement Object.
;                  @Error 0 @Extended 1 Return Object = Success. Returning the created Statement Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $sSQL is called with NULL, a Statement will be created. If you call $sSQL with a SQL string, a Prepared Statement will be created.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLStatementCreate(ByRef $oConnection, $sSQL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oStatement

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSQL) And ($sSQL <> Null) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If IsString($sSQL) Then
		$oStatement = $oConnection.prepareStatement($sSQL)
		If Not IsObj($oStatement) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $oStatement)

	Else
		$oStatement = $oConnection.createStatement()
		If Not IsObj($oStatement) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $oStatement)
	EndIf
EndFunc   ;==>_LOBase_SQLStatementCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLStatementExecuteQuery
; Description ...: Execute a SQL Statement or a Prepared SQL Statement Query.
; Syntax ........: _LOBase_SQLStatementExecuteQuery(ByRef $oStatement[, $sSQL = Null[, $bWritable = False]])
; Parameters ....: $oStatement          - [in/out] an object. A Statement object returned by a previous _LOBase_SQLStatementCreate function.
;                  $sSQL                - [optional] a string value. Default is Null. If the statement being called is not a Prepared Statement, the SQL query will be called here.
;                  $bWritable           - [optional] a boolean value. Default is False. If True, returns a readable and writable Result set. Only works for non-Prepared Statements.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oStatement not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oStatement not a Statement or Prepared Statement Object.
;                  @Error 1 @Extended 3 Return 0 = Statement called in $oStatement is not a Prepared Statement, and $sSQL is not a String.
;                  @Error 1 @Extended 4 Return 0 = $bWritable not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.RowSet" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to perform the Query.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the Result set returned from the SQL Statement Query.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLStatementExecuteQuery(ByRef $oStatement, $sSQL = Null, $bWritable = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResult, $oRowSet, $oServiceManager

	If Not IsObj($oStatement) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oStatement.supportsService("com.sun.star.sdbc.Statement") Or $oStatement.supportsService("com.sun.star.sdbc.PreparedStatement")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSQL) And $oStatement.supportsService("com.sun.star.sdbc.Statement") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bWritable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If $oStatement.supportsService("com.sun.star.sdbc.Statement") Then
		If $bWritable Then
			$oServiceManager = __LO_ServiceManager()
			If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oRowSet = $oServiceManager.createInstance("com.sun.star.sdb.RowSet")
			If Not IsObj($oRowSet) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$oRowSet.activeConnection = $oStatement.Connection()
			$oRowSet.CommandType = $LOB_REP_CONTENT_TYPE_SQL
			$oRowSet.Command = $sSQL
			$oRowSet.Execute()
			$oResult = $oRowSet

		Else
			$oResult = $oStatement.ExecuteQuery($sSQL)
		EndIf

	Else
		$oResult = $oStatement.ExecuteQuery()
	EndIf

	If Not IsObj($oResult) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oResult)
EndFunc   ;==>_LOBase_SQLStatementExecuteQuery

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLStatementExecuteUpdate
; Description ...: Execute a SQL Update Statement or a Prepared SQL Update Statement.
; Syntax ........: _LOBase_SQLStatementExecuteUpdate(ByRef $oStatement[, $sSQL = Null])
; Parameters ....: $oStatement          - [in/out] an object. A Statement object returned by a previous _LOBase_SQLStatementCreate function.
;                  $sSQL                - [optional] a string value. Default is Null. If the statement being called is not a Prepared Statement, the SQL update command will be called here.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oStatement not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oStatement not a Statement or Prepared Statement Object.
;                  @Error 1 @Extended 3 Return 0 = Statement called in $oStatement is not a Prepared Statement, and $sSQL is not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to perform the Update.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning the Row count for INSERT, DELETE or UPDATE SQL Statements, or 0 for SQL Statements that return nothing.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLStatementExecuteUpdate(ByRef $oStatement, $sSQL = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iResult

	If Not IsObj($oStatement) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oStatement.supportsService("com.sun.star.sdbc.Statement") Or $oStatement.supportsService("com.sun.star.sdbc.PreparedStatement")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSQL) And $oStatement.supportsService("com.sun.star.sdbc.Statement") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If $oStatement.supportsService("com.sun.star.sdbc.Statement") Then
		$iResult = $oStatement.ExecuteUpdate($sSQL)

	Else
		$iResult = $oStatement.ExecuteUpdate()
	EndIf

	If Not IsInt($iResult) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iResult)
EndFunc   ;==>_LOBase_SQLStatementExecuteUpdate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_SQLStatementPreparedSetData
; Description ...: Set or clear the current Prepared Statement data.
; Syntax ........: _LOBase_SQLStatementPreparedSetData(ByRef $oStatement[, $iCommand = Null[, $iSetType = Null[, $vValue = Null]]])
; Parameters ....: $oStatement          - [in/out] an object. A Statement object returned by a previous _LOBase_SQLStatementCreate function.
;                  $iCommand            - [optional] an integer value. Default is Null. The command number in the SQL statement to set the data for. 1 based.
;                  $iSetType            - [optional] an integer value (0-16). Default is Null. The type of Set command to perform. See Constants, $LOB_DATA_SET_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  $vValue              - [optional] a variant value. Default is Null. The Data value to set the SQL statement placeholder to.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oStatement not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oStatement not a Statement or Prepared Statement Object.
;                  @Error 1 @Extended 3 Return 0 = $iCommand not an Integer, or less than 1.
;                  @Error 1 @Extended 4 Return 0 = $iSetType not an Integer, less than 0 or greater than 16. See Constants, $LOB_DATA_SET_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iSetType called with $LOB_DATA_SET_TYPE_BOOL and $vValue is not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iSetType called with $LOB_DATA_SET_TYPE_BYTE or $LOB_DATA_SET_TYPE_BYTES and $vValue is not a Binary value.
;                  @Error 1 @Extended 7 Return 0 = $iSetType called with $LOB_DATA_SET_TYPE_SHORT, $LOB_DATA_SET_TYPE_INT, or $LOB_DATA_SET_TYPE_LONG and $vValue is not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $iSetType called with $LOB_DATA_SET_TYPE_FLOAT, or $LOB_DATA_SET_TYPE_DOUBLE and $vValue is not a Number.
;                  @Error 1 @Extended 9 Return 0 = $iSetType called with $LOB_DATA_SET_TYPE_STRING and $vValue is not a String.
;                  @Error 1 @Extended 10 Return 0 = $iSetType called with $LOB_DATA_SET_TYPE_DATE, $LOB_DATA_SET_TYPE_TIME, or $LOB_DATA_SET_TYPE_TIMESTAMP and $vValue is not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.util.Date" Struct.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.util.Time" Struct.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully set the prepared SQL statement Data.
;                  @Error 0 @Extended 1 Return 1 = Success. Successfully cleared the SQL prepared statement of the set data.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with Null in all optional parameters to clear the Prepared Statement placeholders or Data.
;                  When setting a Date, Time, or Timestamp, use a Date Structure returned from the function _LOBase_DateStructCreate, depending on the command, the appropriate values will be copied over to an appropriate structure (Date, or Time) inside the function.
;                  Not all SetTypes have error checking. It is the user's duty to know the values being input are correct.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_SQLStatementPreparedSetData(ByRef $oStatement, $iCommand = Null, $iSetType = Null, $vValue = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tDateTime

	If Not IsObj($oStatement) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not ($oStatement.supportsService("com.sun.star.sdbc.Statement") Or $oStatement.supportsService("com.sun.star.sdbc.PreparedStatement")) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($iCommand, $iSetType, $vValue) Then
		$oStatement.clearParameters()

		Return SetError($__LO_STATUS_SUCCESS, 1, 1)
	EndIf

	If Not __LO_IntIsBetween($iCommand, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iSetType, $LOB_DATA_SET_TYPE_NULL, $LOB_DATA_SET_TYPE_OBJECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	Switch $iSetType
		Case $LOB_DATA_SET_TYPE_NULL
			$oStatement.setNull($iCommand, $LOB_DATA_TYPE_SQLNULL)

		Case $LOB_DATA_SET_TYPE_BOOL
			If Not IsBool($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oStatement.setBoolean($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_BYTE
			If Not IsBinary($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oStatement.setByte($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_BYTES
			If Not IsBinary($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oStatement.setBytes($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_SHORT
			If Not IsInt($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oStatement.setShort($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_INT
			If Not IsInt($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oStatement.setInt($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_LONG
			If Not IsInt($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$oStatement.setLong($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_FLOAT
			If Not IsNumber($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oStatement.setFloat($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_DOUBLE
			If Not IsNumber($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$oStatement.setDouble($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_STRING
			If Not IsString($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$oStatement.setString($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_DATE
			If Not IsObj($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			$tDateTime = __LO_CreateStruct("com.sun.star.util.Date")
			If Not IsObj($tDateTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$tDateTime.Year = $vValue.Year()
			$tDateTime.Month = $vValue.Month()
			$tDateTime.Day = $vValue.Day()

			$oStatement.setDate($iCommand, $tDateTime)

		Case $LOB_DATA_SET_TYPE_TIME
			If Not IsObj($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			$tDateTime = __LO_CreateStruct("com.sun.star.util.Time")
			If Not IsObj($tDateTime) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$tDateTime.Hours = $vValue.Hours()
			$tDateTime.Minutes = $vValue.Minutes()
			$tDateTime.Seconds = $vValue.Seconds()
			$tDateTime.NanoSeconds = $vValue.NanoSeconds()
			If __LO_VersionCheck(4.1) Then $tDateTime.IsUTC = $vValue.IsUTC()

			$oStatement.setTime($iCommand, $tDateTime)

		Case $LOB_DATA_SET_TYPE_TIMESTAMP
			If Not IsObj($vValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			$oStatement.setTimestamp($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_ARRAY
			$oStatement.setArray($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_CLOB
			$oStatement.setClob($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_BLOB
			$oStatement.setBlob($iCommand, $vValue)

		Case $LOB_DATA_SET_TYPE_OBJECT
			$oStatement.setObject($iCommand, $vValue)
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_SQLStatementPreparedSetData
