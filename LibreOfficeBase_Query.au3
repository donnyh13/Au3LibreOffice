#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

;~ #Tidy_Parameters=/sf /reel
#include-once
#include "LibreOffice_Constants.au3"

; Common includes for Base
#include "LibreOfficeBase_Internal.au3"

; Other includes for Base

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Adding, Deleting, and modifying, etc. L.O. Base Queries.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_QueriesGetCount
; _LOBase_QueriesGetNames
; _LOBase_QueryAddByName
; _LOBase_QueryAddBySQL
; _LOBase_QueryDelete
; _LOBase_QueryExists
; _LOBase_QueryFieldGetObjByIndex
; _LOBase_QueryFieldGetObjByName
; _LOBase_QueryFieldModify
; _LOBase_QueryFieldsGetCount
; _LOBase_QueryFieldsGetNames
; _LOBase_QueryGetObjByIndex
; _LOBase_QueryGetObjByName
; _LOBase_QueryName
; _LOBase_QuerySQLCommand
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueriesGetCount
; Description ...: Retrieve a count of Queries contained in the Document.
; Syntax ........: _LOBase_QueriesGetCount(ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve count of Queries.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Queries contained in the Document as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueriesGetCount(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oConnection.ImplementationName() <> "com.sun.star.sdbc.drivers.OConnectionWrapper" Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oConnection.Queries.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOBase_QueriesGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueriesGetNames
; Description ...: Retrieve an Array of Query Names contained in the Document.
; Syntax ........: _LOBase_QueriesGetNames(ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Array of Element names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Query names contained in this Document. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueriesGetNames(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oConnection.ImplementationName() <> "com.sun.star.sdbc.drivers.OConnectionWrapper" Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asNames = $oConnection.Queries.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOBase_QueriesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryAddByName
; Description ...: Add a Query to a Database by Name.
; Syntax ........: _LOBase_QueryAddByName(ByRef $oConnection, $sQueryName, $sSourceName, $sFieldName)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sQueryName          - a string value. The Unique name of the Query to create.
;                  $sSourceName         - a string value. The Table or Query Name to use as a Source.
;                  $sFieldName          - a string value. The Field name to reference from the Table or Query called in $sSourceName. Accepts "*" also.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $sQueryName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sSourceName not a String.
;                  @Error 1 @Extended 5 Return 0 = $sFieldName not a String.
;                  @Error 1 @Extended 6 Return 0 = Document already contains a Query with the name called in $sQueryName.
;                  @Error 1 @Extended 7 Return 0 = Document already contains a Table with the name called in $sQueryName.
;                  @Error 1 @Extended 8 Return 0 = Query or Table with name called in $sSourceName not found.
;                  @Error 1 @Extended 9 Return 0 = Source called in $sSourceName does not contain a field with name as called in $sFieldName.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Query Descriptor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Queries Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Source Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Database specific Quotation character.
;                  @Error 3 @Extended 5 Return 0 = Failed to insert new Query.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve New Query's Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning new Query's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sQueryName must be unique from both Query and Table names.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryAddByName(ByRef $oConnection, $sQueryName, $sSourceName, $sFieldName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQuery, $oQueries, $oQueryDesc, $oSource
	Local $sQuote

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oConnection.ImplementationName() <> "com.sun.star.sdbc.drivers.OConnectionWrapper" Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sQueryName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSourceName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sFieldName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oQueries = $oConnection.Queries()
	If Not IsObj($oQueries) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If $oQueries.hasByName($sQueryName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If $oConnection.Tables.hasByName($sQueryName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If (Not $oQueries.hasByName($sSourceName) And Not $oConnection.Tables.hasByName($sSourceName)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	If $oQueries.hasByName($sSourceName) Then
		$oSource = $oQueries.getByName($sSourceName)
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	ElseIf $oConnection.Tables.hasByName($sSourceName) Then
		$oSource = $oConnection.Tables.getByName($sSourceName)
		If Not IsObj($oSource) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	EndIf

	$sQuote = $oConnection.MetaData.getIdentifierQuoteString()
	If Not IsString($sQuote) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	If ($sFieldName <> "*") And Not $oSource.Columns.hasByName($sFieldName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	$oQueryDesc = $oQueries.createDataDescriptor()
	If Not IsObj($oQueryDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oQueryDesc.Name = $sQueryName

	If $sFieldName <> "*" Then
		$oQueryDesc.Command = "SELECT " & $sQuote & $sFieldName & $sQuote & " FROM " & $sQuote & $sSourceName & $sQuote

	Else
		$oQueryDesc.Command = "SELECT " & $sFieldName & " FROM " & $sQuote & $sSourceName & $sQuote
	EndIf

	$oQueries.appendByDescriptor($oQueryDesc)

	If Not $oQueries.hasByName($sQueryName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$oQuery = $oQueries.getByName($sQueryName)
	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oQuery)
EndFunc   ;==>_LOBase_QueryAddByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryAddBySQL
; Description ...: Add a Query to a Database using an SQL Command.
; Syntax ........: _LOBase_QueryAddBySQL(ByRef $oConnection, $sQueryName, $sSQL_Command)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sQueryName          - a string value. The Unique name of the Query to create.
;                  $sSQL_Command        - a string value. The SQL Query Command to initialize the new Query with.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $sQueryName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sSQL_Command not a String.
;                  @Error 1 @Extended 5 Return 0 = Document already contains a Query with the name called in $sQueryName.
;                  @Error 1 @Extended 6 Return 0 = Document already contains a Table with the name called in $sQueryName.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Query Descriptor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Queries Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to insert new Query.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve New Query's Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning new Query's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: It is the user's responsibility to ensure Table, Query, and Field names called in the SQL command are correct.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryAddBySQL(ByRef $oConnection, $sQueryName, $sSQL_Command)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQuery, $oQueries, $oQueryDesc

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oConnection.ImplementationName() <> "com.sun.star.sdbc.drivers.OConnectionWrapper" Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sQueryName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sSQL_Command) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oQueries = $oConnection.Queries()
	If Not IsObj($oQueries) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If $oQueries.hasByName($sQueryName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If $oConnection.Tables.hasByName($sQueryName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oQueryDesc = $oQueries.createDataDescriptor()
	If Not IsObj($oQueryDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oQueryDesc.Name = $sQueryName

	$oQueryDesc.Command = $sSQL_Command

	$oQueries.appendByDescriptor($oQueryDesc)

	If Not $oQueries.hasByName($sQueryName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oQuery = $oQueries.getByName($sQueryName)
	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oQuery)
EndFunc   ;==>_LOBase_QueryAddBySQL

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryDelete
; Description ...: Delete a Query from the Document.
; Syntax ........: _LOBase_QueryDelete(ByRef $oConnection, ByRef $oQuery)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $oQuery              - [in/out] an object. A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $oQuery not an Object.
;                  @Error 1 @Extended 4 Return 0 = Connection called in $oConnection does not contain the Query called in $oQuery.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Queries Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Query name.
;                  @Error 3 @Extended 4 Return 0 = Failed to delete Query.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Query was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryDelete(ByRef $oConnection, ByRef $oQuery)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQueries
	Local $sName

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oConnection.ImplementationName() <> "com.sun.star.sdbc.drivers.OConnectionWrapper" Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oQueries = $oConnection.Queries()
	If Not IsObj($oQueries) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$sName = $oQuery.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If Not $oQueries.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oQueries.dropByName($sName)

	If $oQueries.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_QueryDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryExists
; Description ...: Check whether a Document contains a Query by name.
; Syntax ........: _LOBase_QueryExists(ByRef $oConnection, $sName)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sName               - a string value. The name of the Query to look for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Queries Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to query Queries for Query name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning a Boolean value indicating if the Document contains a Query by the called name (True) or not.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryExists(ByRef $oConnection, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQueries
	Local $bReturn

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oConnection.ImplementationName() <> "com.sun.star.sdbc.drivers.OConnectionWrapper" Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oQueries = $oConnection.Queries()
	If Not IsObj($oQueries) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$bReturn = $oQueries.hasByName($sName)
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_QueryExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryFieldGetObjByIndex
; Description ...: Retrieve a Query Field's Object by Index.
; Syntax ........: _LOBase_QueryFieldGetObjByIndex(ByRef $oQuery, $iField)
; Parameters ....: $oQuery              - [in/out] an object. A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
;                  $iField              - an integer value. The Index value of the Field to retrieve the Object for. 0 Based.
; Return values . Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oQuery not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iField not an Integer, less than 0, or greater than number of Fields contained in the query.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Column Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Column's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryFieldGetObjByIndex(ByRef $oQuery, $iField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumn, $oColumns

	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oQuery.Columns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If Not __LOBase_IntIsBetween($iField, 0, $oColumns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumn = $oColumns.getByIndex($iField)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOBase_QueryFieldGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryFieldGetObjByName
; Description ...: Retrieve a Query Field's Object by name.
; Syntax ........: _LOBase_QueryFieldGetObjByName(ByRef $oQuery, $sName)
; Parameters ....: $oQuery              - [in/out] an object. A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
;                  $sName               - a string value. The Query Field name to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oQuery not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Query does not contain a Field with the name called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Column Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Column's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Field name called in $sName must be the Alias name, if present, otherwise the real name will work.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryFieldGetObjByName(ByRef $oQuery, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumn, $oColumns

	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumns = $oQuery.Columns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If Not $oColumns.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumn = $oColumns.getByName($sName)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOBase_QueryFieldGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryFieldModify
; Description ...: Set or Retrieve Query Field settings.
; Syntax ........: _LOBase_QueryFieldModify(ByRef $oField[, $sAlias = Null[, $bVisible = Null[, $sRealName = Null]]])
; Parameters ....: $oField              - [in/out] an object. A Query field object returned by a previous _LOBase_QueryFieldGetObjByIndex, or _LOBase_QueryFieldGetObjByName function.
;                  $sAlias              - [optional] a string value. Default is Null. The Alias to call the present field in this Query.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the Query Field will be visible in the Query results.
;                  $sRealName           - [optional] a string value. Default is Null. This parameter is not settable, but indicates in what position the Field's real name will be returned.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oField not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sAlias not a String.
;                  @Error 1 @Extended 3 Return 0 = $bVisible not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sAlias
;                  |                               2 = Error setting $bVisible
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sRealName modifies nothing, but is an indicator of where the Query Field's Real name (The name without an Alias) will be returned when returning the current settings.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryFieldModify(ByRef $oField, $sAlias = Null, $bVisible = Null, $sRealName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSettings[3]

	If Not IsObj($oField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOBase_VarsAreNull($sAlias, $bVisible, $sRealName) Then
		__LOBase_ArrayFill($avSettings, $oField.Name(), $oField.Hidden(), $oField.RealName())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avSettings)
	EndIf

	If ($sAlias <> Null) Then
		If Not IsString($sAlias) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oField.setName($sAlias)
		$iError = ($oField.Name() = $sAlias) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bVisible <> Null) Then
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oField.Hidden = $bVisible
		$iError = ($oField.Hidden() = $bVisible) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_QueryFieldModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryFieldsGetCount
; Description ...: Retrieve a count of Fields referenced in a Query.
; Syntax ........: _LOBase_QueryFieldsGetCount(ByRef $oQuery)
; Parameters ....: $oQuery              - [in/out] an object. A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oQuery not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve count of Queries.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Queries contained in the document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryFieldsGetCount(ByRef $oQuery)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oQuery.Columns.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOBase_QueryFieldsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryFieldsGetNames
; Description ...: Retrieve an Array of Fields referenced in a Query.
; Syntax ........: _LOBase_QueryFieldsGetNames(ByRef $oQuery)
; Parameters ....: $oQuery              - [in/out] an object. A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oQuery not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Array of Query names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning array of Query names. @Extended will be set to the number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The name returned will be the Alias of the field, if there is one.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryFieldsGetNames(ByRef $oQuery)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]

	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asNames = $oQuery.Columns.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOBase_QueryFieldsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryGetObjByIndex
; Description ...: Retrieve a Query's Object by Index.
; Syntax ........: _LOBase_QueryGetObjByIndex(ByRef $oConnection, $iQuery)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $iQuery              - an integer value. The Index value of the Query to retrieve. 0 Based.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $iQuery not an Integer, less than 0, or greater than number of Queries contained in the Database.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Queries Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Query Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Query's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryGetObjByIndex(ByRef $oConnection, $iQuery)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQuery, $oQueries

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oConnection.ImplementationName() <> "com.sun.star.sdbc.drivers.OConnectionWrapper" Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oQueries = $oConnection.Queries()
	If Not IsObj($oQueries) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not __LOBase_IntIsBetween($iQuery, 0, $oQueries.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oQuery = $oQueries.getByIndex($iQuery)
	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oQuery)
EndFunc   ;==>_LOBase_QueryGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryGetObjByName
; Description ...: Retrieve a Query's Object by name.
; Syntax ........: _LOBase_QueryGetObjByName(ByRef $oConnection, $sName)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sName               - a string value. The Query's name to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = Query with name called in $sName not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Queries Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Query Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Query's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryGetObjByName(ByRef $oConnection, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQuery, $oQueries

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oConnection.ImplementationName() <> "com.sun.star.sdbc.drivers.OConnectionWrapper" Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oQueries = $oConnection.Queries()
	If Not IsObj($oQueries) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not $oQueries.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oQuery = $oQueries.getByName($sName)
	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oQuery)
EndFunc   ;==>_LOBase_QueryGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryName
; Description ...: Set or Retrieve the Query's name.
; Syntax ........: _LOBase_QueryName(ByRef $oQuery[, $sName = Null])
; Parameters ....: $oQuery              - [in/out] an object. A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
;                  $sName               - [optional] a string value. Default is Null. The new name to set the Query to. See Remarks.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oQuery not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. $sName set to Null, returning current Query Name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function does not check if the new name already exists in Tables or Queries.
;                  According to LibreOffice SDK API IDL XRename Interface, It would seem some Database types don't support the renaming of Queries.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryName(ByRef $oQuery, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($sName = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oQuery.Name())

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oQuery.rename($sName)
	If ($oQuery.Name() <> $sName) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_QueryName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QuerySQLCommand
; Description ...: Set or Retrieve the Query's SQL command.
; Syntax ........: _LOBase_QuerySQLCommand(ByRef $oQuery[, $sSQL_Command = Null])
; Parameters ....: $oQuery              - [in/out] an object. A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
;                  $sSQL_Command        - [optional] a string value. Default is Null. The SQL command to set for the Query.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oQuery not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sSQL_Command not a String.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sSQL_Command
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. $sSQL_Command set to Null, returning current Query SQL Command.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QuerySQLCommand(ByRef $oQuery, $sSQL_Command = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($sSQL_Command = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oQuery.Command())

	If Not IsString($sSQL_Command) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oQuery.Command = $sSQL_Command
	If ($oQuery.Command() <> $sSQL_Command) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_QuerySQLCommand
