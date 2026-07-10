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
; _LOBase_QueryDocClose
; _LOBase_QueryDocConnect
; _LOBase_QueryDocGetName
; _LOBase_QueryDocGetRowSet
; _LOBase_QueryDocOpenByName
; _LOBase_QueryDocOpenByObject
; _LOBase_QueryDocVisible
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
; Parameters ....: $oConnection         - A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning count of Queries contained in the Document as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oConnection not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oConnection not a Connection Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Connection called in $oConnection is closed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve count of Queries.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueriesGetNames, _LOBase_QueryGetObjByIndex
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueriesGetCount(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oConnection.Queries.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOBase_QueriesGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueriesGetNames
; Description ...: Retrieve an Array of Query Names contained in the Document.
; Syntax ........: _LOBase_QueriesGetNames(ByRef $oConnection)
; Parameters ....: $oConnection         - A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning Array of Query names contained in this Document. @Extended is set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oConnection not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oConnection not a Connection Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Connection called in $oConnection is closed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Array of Element names.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueriesGetCount, _LOBase_QueryDocOpenByName, _LOBase_QueryDelete, _LOBase_QueryExists, _LOBase_QueryGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueriesGetNames(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asNames = $oConnection.Queries.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOBase_QueriesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryAddByName
; Description ...: Add a Query to a Database by Name.
; Syntax ........: _LOBase_QueryAddByName(ByRef $oConnection, $sQueryName, $sSourceName, $sFieldName)
; Parameters ....: $oConnection         - A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sQueryName          - The Unique name of the Query to create.
;                  $sSourceName         - The Table or Query Name to use as a Source.
;                  $sFieldName          - The Field name to reference from the Table or Query called in $sSourceName. Accepts "*" also.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning new Query's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oConnection not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oConnection not a Connection Object.
;                  @Error: 1, @Extended: 3 = $sQueryName not a String.
;                  @Error: 1, @Extended: 4 = $sSourceName not a String.
;                  @Error: 1, @Extended: 5 = $sFieldName not a String.
;                  @Error: 1, @Extended: 6 = Document already contains a Query with the name called in $sQueryName.
;                  @Error: 1, @Extended: 7 = Document already contains a Table with the name called in $sQueryName.
;                  @Error: 1, @Extended: 8 = Query or Table with name called in $sSourceName not found.
;                  @Error: 1, @Extended: 9 = Source called in $sSourceName does not contain a field with name as called in $sFieldName.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a Query Descriptor.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Connection called in $oConnection is closed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Queries Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Source Object.
;                  @Error: 3, @Extended: 4 = Failed to retrieve Database specific Quotation character.
;                  @Error: 3, @Extended: 5 = Failed to insert new Query.
;                  @Error: 3, @Extended: 6 = Failed to retrieve New Query's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sQueryName must be unique from both Query and Table names.
; Related .......: _LOBase_QueryAddBySQL, _LOBase_QueryExists, _LOBase_QueryDocOpenByObject, _LOBase_QuerySQLCommand
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryAddByName(ByRef $oConnection, $sQueryName, $sSourceName, $sFieldName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQuery, $oQueries, $oQueryDesc, $oSource
	Local $sQuote

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
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
; Parameters ....: $oConnection         - A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sQueryName          - The Unique name of the Query to create.
;                  $sSQL_Command        - The SQL Query Command to initialize the new Query with.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning new Query's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oConnection not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oConnection not a Connection Object.
;                  @Error: 1, @Extended: 3 = $sQueryName not a String.
;                  @Error: 1, @Extended: 4 = $sSQL_Command not a String.
;                  @Error: 1, @Extended: 5 = Document already contains a Query with the name called in $sQueryName.
;                  @Error: 1, @Extended: 6 = Document already contains a Table with the name called in $sQueryName.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a Query Descriptor.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Connection called in $oConnection is closed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Queries Object.
;                  @Error: 3, @Extended: 3 = Failed to insert new Query.
;                  @Error: 3, @Extended: 4 = Failed to retrieve New Query's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: It is the user's responsibility to ensure Table, Query, and Field names called in the SQL command are correct.
; Related .......: _LOBase_QueryAddByName, _LOBase_QueryExists, _LOBase_QueryDocOpenByObject, _LOBase_QuerySQLCommand
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryAddBySQL(ByRef $oConnection, $sQueryName, $sSQL_Command)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQuery, $oQueries, $oQueryDesc

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
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
; Parameters ....: $oConnection         - A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $oQuery              - A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Query was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oConnection not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oConnection not a Connection Object.
;                  @Error: 1, @Extended: 3 = $oQuery not an Object.
;                  @Error: 1, @Extended: 4 = Connection called in $oConnection does not contain the Query called in $oQuery.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Connection called in $oConnection is closed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Queries Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Query name.
;                  @Error: 3, @Extended: 4 = Failed to delete Query.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueriesGetNames, _LOBase_QueryAddByName, _LOBase_QueryAddBySQL, _LOBase_QueryExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryDelete(ByRef $oConnection, ByRef $oQuery)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQueries
	Local $sName

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oQueries = $oConnection.Queries()
	If Not IsObj($oQueries) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$sName = $oQuery.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If Not $oQueries.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oQueries.dropByName($sName)

	If $oQueries.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$oQuery = Null

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_QueryDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryDocClose
; Description ...: Close a Query Document.
; Syntax ........: _LOBase_QueryDocClose(ByRef $oQueryDoc[, $bDeliverOwnership = True])
; Parameters ....: $oQueryDoc           - A Query Document Object from a previous _LOBase_QueryDocOpenByName, _LOBase_QueryDocOpenByObject or _LOBase_QueryDocConnect function.
;                  $bDeliverOwnership   - [optional] Default is True. If True, deliver ownership of the Query Document Object from the script to LibreOffice, recommended is True.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully closed the Query Document.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oQueryDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bDeliverOwnership not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to close the Query Document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueryDocOpenByName, _LOBase_QueryDocOpenByObject, _LOBase_QueryDocConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryDocClose(ByRef $oQueryDoc, $bDeliverOwnership = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oQueryDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bDeliverOwnership) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oQueryDoc.Frame.close($bDeliverOwnership)

	If Not __LO_IsObjInvalid($oQueryDoc, "ComponentWindow.Windows") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oQueryDoc = Null

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_QueryDocClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryDocConnect
; Description ...: Connect to an open instance of a Database Query Document.
; Syntax ........: _LOBase_QueryDocConnect([$iMode = $LO_DOC_CONNECT_MODE_CURRENT])
; Parameters ....: $iMode               - [optional] (0-1) Default is $LO_DOC_CONNECT_MODE_CURRENT. The Connect mode. See Constants, $LO_DOC_CONNECT_MODE_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: Object or Array.
;                  @Error: 0, @Extended: ?, Return: Object = Success, The Object for the current, or last active Base Query document is returned. @Extended set to Document type Constant as an Integer. See Constants, $LO_DOC_TYPE_* as defined in LibreOffice_Constants.au3.
;                  @Error: 0, @Extended: ?, Return: Array = Success, A two columned Array of all open LibreOffice Base Query Documents. @Extended is set to number of results. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $iMode not an Integer, less than 0 or greater than 1. See Constants, $LO_DOC_CONNECT_MODE_* as defined in LibreOffice_Constants.au3.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating ServiceManager object.
;                  @Error: 2, @Extended: 2 = Error creating Desktop object.
;                  @Error: 2, @Extended: 3 = Error creating enumeration of open documents.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = No open LibreOffice documents.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Document Object.
;                  @Error: 3, @Extended: 3 = Failed to identify Document type.
;                  @Error: 3, @Extended: 4 = Current Document not a Base Query Document.
;                  @Error: 3, @Extended: 5 = No matches found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only Base Query documents are returned using either of the flags.
;                  The Connect All option returns a two columned array. ($aArray[0][2]), each result is stored in a separate row.
;                  -Row 1 Column 0 contains the Object for that document. e.g. $aArray[0][0] = $oDoc
;                  -Row 1 Column 1 contains the Document's type Constant as an Integer. See Constants, $LO_DOC_TYPE_* as defined in LibreOffice_Constants.au3. e.g.: $aArray[0][1] = $LO_DOC_TYPE_BASE_FORM_VIEW.
;                  -Row 2 contains the Object for the next document. e.g. $aArray[1][0] = $oDoc2. And so on.
; Related .......: _LOBase_QueryDocOpenByName, _LOBase_QueryDocOpenByObject, _LOBase_QueryDocClose
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryDocConnect($iMode = $LO_DOC_CONNECT_MODE_CURRENT)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0, $iDocType
	Local $aoConnectAll[0][2]
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop

	If Not __LO_IntIsBetween($iMode, $LO_DOC_CONNECT_MODE_ALL, $LO_DOC_CONNECT_MODE_CURRENT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

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

				If __LO_IntIsBetween($iDocType, $LO_DOC_TYPE_BASE_QUERY_DESIGN, $LO_DOC_TYPE_BASE_QUERY_VIEW) Then
					If (UBound($aoConnectAll) <= $iCount) Then ReDim $aoConnectAll[$iCount + 1][2]
					$aoConnectAll[$iCount][0] = $oDoc
					$aoConnectAll[$iCount][1] = $iDocType
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
			If Not __LO_IntIsBetween($iDocType, $LO_DOC_TYPE_BASE_QUERY_DESIGN, $LO_DOC_TYPE_BASE_QUERY_VIEW) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; Not a Base Form Doc.

			Return SetError($__LO_STATUS_SUCCESS, $iDocType, $oDoc)
	EndSwitch

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; No matches
EndFunc   ;==>_LOBase_QueryDocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryDocGetName
; Description ...: Retrieve the Query document's name.
; Syntax ........: _LOBase_QueryDocGetName(ByRef $oQueryDoc[, $bReturnFull = False])
; Parameters ....: $oQueryDoc           - A Query Document Object from a previous _LOBase_QueryDocOpenByName, _LOBase_QueryDocOpenByObject or _LOBase_QueryDocConnect function.
;                  $bReturnFull         - [optional] Default is False. If True, the full window title is returned, such as is used by AutoIt window related functions.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Success. Returning the document's Name as a String. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oQueryDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bReturnFull not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Document's name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bReturnFull is True, the return value will be one of the following:
;                  If the Query Document is in Design mode: "<Database Doc name>.<extension> : <Query name> — LibreOffice Base: Query Design" e.g. "Testing.odb : QryAutoIt — LibreOffice Base: Query Design".
;                  If the Query Document is in Viewing mode: "<Query name> - <Database Doc name> — LibreOffice Base: Table Data View" e.g. "QryAutoIt - Testing — LibreOffice Base: Table Data View"
;                  Else if $bReturnFull is False, the return value will be one of the following:
;                  If the Query Document is in Design mode: "<Database Doc name>.<extension> : <Query name>", e.g. "Testing.odb : QryAutoIt"
;                  If the Query Document is in Viewing mode: "<Query name> - <Database Doc name>", e.g. "QryAutoIt - Testing"
; Related .......: _LOBase_QueryDocClose, _LOBase_QueryName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryDocGetName(ByRef $oQueryDoc, $bReturnFull = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oQueryDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnFull) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $bReturnFull Then
		$sName = $oQueryDoc.Frame.Title()
		If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else
		$sName = $oQueryDoc.Title()
		If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $sName)
EndFunc   ;==>_LOBase_QueryDocGetName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryDocGetRowSet
; Description ...: Retrieve a Row Set for a Query opened for Data entry/Viewing.
; Syntax ........: _LOBase_QueryDocGetRowSet(ByRef $oQueryDoc)
; Parameters ....: $oQueryDoc           - A Query Document Object from a previous _LOBase_QueryDocOpenByName, _LOBase_QueryDocOpenByObject or _LOBase_QueryDocConnect function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning Query's RowSet Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oQueryDoc not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oQueryDoc not Query opened in viewing/data entry mode.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve RowSet Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving the RowSet for the Query allows you to manipulate data contained in the Query using _LOBase_SQLResultRowUpdate, etc. functions.
; Related .......: _LOBase_SQLResultCursorMove, _LOBase_SQLResultCursorQuery, _LOBase_SQLResultRowModify, _LOBase_SQLResultRowQuery, _LOBase_SQLResultRowRead, _LOBase_SQLResultRowRefresh, _LOBase_SQLResultRowUpdate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryDocGetRowSet(ByRef $oQueryDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResultSet

	If Not IsObj($oQueryDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oQueryDoc.supportsService("com.sun.star.sdb.DataSourceBrowser") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oResultSet = $oQueryDoc.FormOperations.Cursor()
	If Not IsObj($oResultSet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oResultSet)
EndFunc   ;==>_LOBase_QueryDocGetRowSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryDocOpenByName
; Description ...: Open a Query Document either in design mode or viewing mode.
; Syntax ........: _LOBase_QueryDocOpenByName(ByRef $oConnection, $sQuery[, $bEdit = False[, $bHidden = False]])
; Parameters ....: $oConnection         - A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sQuery              - The Query's name.
;                  $bEdit               - [optional] Default is False. If True, the Query is opened in editing mode to add or remove columns. If False, the Query is opened in data viewing mode, to modify Query Data.
;                  $bHidden             - [optional] Default is False. If True, the Document will be invisible.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully opened the Query Document, returning its object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oConnection not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oConnection not a Connection Object.
;                  @Error: 1, @Extended: 3 = $sQuery not a String.
;                  @Error: 1, @Extended: 4 = $bEdit not a Boolean.
;                  @Error: 1, @Extended: 5 = $bHidden not a Boolean.
;                  @Error: 1, @Extended: 6 = No Query with name called in $sQuery found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Connection called in $oConnection is closed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Queries Object.
;                  @Error: 3, @Extended: 3 = Failed to create a Connection to Database.
;                  @Error: 3, @Extended: 4 = Failed to open Query Document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueryDocOpenByObject, _LOBase_QueryDocConnect, _LOBase_QueryDocClose, _LOBase_QueryDocVisible
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryDocOpenByName(ByRef $oConnection, $sQuery, $bEdit = False, $bHidden = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQueries, $oQueryDoc
	Local $aArgs[1]

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bEdit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oQueries = $oConnection.getQueries()
	If Not IsObj($oQueries) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oQueries.hasByName($sQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If Not $oConnection.Parent.DatabaseDocument.CurrentController.isConnected() Then $oConnection.Parent.DatabaseDocument.CurrentController.connect()
	If Not $oConnection.Parent.DatabaseDocument.CurrentController.isConnected() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$aArgs[0] = __LO_SetPropertyValue("Hidden", $bHidden)

	$oQueryDoc = $oConnection.Parent.DatabaseDocument.CurrentController.loadComponentWithArguments($LOB_SUB_COMP_TYPE_QUERY, $sQuery, $bEdit, $aArgs)
	If Not IsObj($oQueryDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oQueryDoc)
EndFunc   ;==>_LOBase_QueryDocOpenByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryDocOpenByObject
; Description ...: Open a Query Document either in design mode or viewing mode.
; Syntax ........: _LOBase_QueryDocOpenByObject(ByRef $oConnection, ByRef $oQuery[, $bEdit = False[, $bHidden = False]])
; Parameters ....: $oConnection         - A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $oQuery              - A Query object returned by a previous _LOBase_QueryGetObjByIndex, _LOBase_QueryGetObjByName, _LOBase_QueryAddByName or _LOBase_QueryAddBySQL function.
;                  $bEdit               - [optional] Default is False. If True, the Query is opened in editing mode to add or remove columns. If False, the Query is opened in data viewing mode, to modify Query Data.
;                  $bHidden             - [optional] Default is False. If True, the Document will be invisible.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully opened Query Document, returning its object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oConnection not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oConnection not a Connection Object.
;                  @Error: 1, @Extended: 3 = $oQuery not an Object.
;                  @Error: 1, @Extended: 4 = $bEdit not a Boolean.
;                  @Error: 1, @Extended: 5 = $bHidden not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Connection called in $oConnection is closed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Query Name.
;                  @Error: 3, @Extended: 3 = Failed to create a Connection to Database.
;                  @Error: 3, @Extended: 4 = Failed to open Query Document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueryDocOpenByName, _LOBase_QueryDocConnect, _LOBase_QueryDocClose, _LOBase_QueryDocVisible
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryDocOpenByObject(ByRef $oConnection, ByRef $oQuery, $bEdit = False, $bHidden = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQueryDoc
	Local $sQuery
	Local $aArgs[1]

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bEdit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sQuery = $oQuery.Name()
	If Not IsString($sQuery) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not $oConnection.Parent.DatabaseDocument.CurrentController.isConnected() Then $oConnection.Parent.DatabaseDocument.CurrentController.connect()
	If Not $oConnection.Parent.DatabaseDocument.CurrentController.isConnected() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$aArgs[0] = __LO_SetPropertyValue("Hidden", $bHidden)

	$oQueryDoc = $oConnection.Parent.DatabaseDocument.CurrentController.loadComponentWithArguments($LOB_SUB_COMP_TYPE_QUERY, $sQuery, $bEdit, $aArgs)
	If Not IsObj($oQueryDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oQueryDoc)
EndFunc   ;==>_LOBase_QueryDocOpenByObject

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryDocVisible
; Description ...: Set or Retrieve Query Document Visibility.
; Syntax ........: _LOBase_QueryDocVisible(ByRef $oQueryDoc[, $bVisible = Null])
; Parameters ....: $oQueryDoc           - A Query Document Object from a previous _LOBase_QueryDocOpenByName, _LOBase_QueryDocOpenByObject or _LOBase_QueryDocConnect function.
;                  $bVisible            - [optional] Default is Null. If True, the Query Document will be visible.
; Return values .: Success: 1 or Boolean.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Boolean = Success. All optional parameters were called with Null, returning current visibility setting.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oQueryDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bVisible not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current visibility setting.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......:  _LOBase_QueryDocOpenByName, _LOBase_QueryDocOpenByObject
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryDocVisible(ByRef $oQueryDoc, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oQueryDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bVisible) Then
		$bVisible = $oQueryDoc.Frame.ContainerWindow.IsVisible()
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $bVisible)
	EndIf

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oQueryDoc.Frame.ContainerWindow.Visible = $bVisible
	$iError = ($oQueryDoc.Frame.ContainerWindow.IsVisible() = $bVisible) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_QueryDocVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryExists
; Description ...: Check whether a Document contains a Query by name.
; Syntax ........: _LOBase_QueryExists(ByRef $oConnection, $sName)
; Parameters ....: $oConnection         - A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sName               - The name of the Query to look for.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning a Boolean value indicating if the Document contains a Query by the called name (True) or not.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oConnection not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oConnection not a Connection Object.
;                  @Error: 1, @Extended: 3 = $sName not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Connection called in $oConnection is closed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Queries Object.
;                  @Error: 3, @Extended: 3 = Failed to query Queries for Query name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueryDelete, _LOBase_QueryDocOpenByName, _LOBase_QueriesGetNames, _LOBase_QueryGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryExists(ByRef $oConnection, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQueries
	Local $bReturn

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
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
; Parameters ....: $oQuery              - A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
;                  $iField              - The Index value of the Field to retrieve the Object for. 0 Based.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested Column's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oQuery not an Object.
;                  @Error: 1, @Extended: 2 = $iField not an Integer, less than 0 or greater than number of Fields contained in the query.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Columns Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Column Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueryFieldGetObjByName, _LOBase_QueryFieldsGetCount, _LOBase_QueryFieldModify, _LOBase_QueryGetObjByIndex, _LOBase_QueryGetObjByName
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
	If Not __LO_IntIsBetween($iField, 0, $oColumns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumn = $oColumns.getByIndex($iField)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOBase_QueryFieldGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryFieldGetObjByName
; Description ...: Retrieve a Query Field's Object by name.
; Syntax ........: _LOBase_QueryFieldGetObjByName(ByRef $oQuery, $sName)
; Parameters ....: $oQuery              - A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
;                  $sName               - The Query Field name to retrieve the Object for.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested Column's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oQuery not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  @Error: 1, @Extended: 3 = Query does not contain a Field with the name called in $sName.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Columns Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Column Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Field name called in $sName must be the Alias name, if present, otherwise the real name will work.
; Related .......: _LOBase_QueryFieldsGetNames, _LOBase_QueryFieldGetObjByIndex, _LOBase_QueryFieldModify, _LOBase_QueryGetObjByIndex, _LOBase_QueryGetObjByName
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
; Parameters ....: $oField              - A Query field object returned by a previous _LOBase_QueryFieldGetObjByIndex, or _LOBase_QueryFieldGetObjByName function.
;                  $sAlias              - [optional] Default is Null. The Alias to call the present field in this Query.
;                  $bVisible            - [optional] Default is Null. If True, the Query Field will be visible in the Query results.
;                  $sRealName           - [optional] Default is Null. This parameter is not settable, but indicates in what position the Field's real name will be returned.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oField not an Object.
;                  @Error: 1, @Extended: 2 = $sAlias not a String.
;                  @Error: 1, @Extended: 3 = $bVisible not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sAlias
;                  |                               2 = Error setting $bVisible
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $sRealName modifies nothing, but is an indicator of where the Query Field's Real name (The name without an Alias) will be returned when returning the current settings.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOBase_QueryFieldGetObjByIndex, _LOBase_QueryFieldGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryFieldModify(ByRef $oField, $sAlias = Null, $bVisible = Null, $sRealName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSettings[3]

	If Not IsObj($oField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sAlias, $bVisible, $sRealName) Then
		__LO_ArrayFill($avSettings, $oField.Name(), $oField.Hidden(), $oField.RealName())

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
; Parameters ....: $oQuery              - A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning count of Queries contained in the document.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oQuery not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve count of Queries.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueryFieldGetObjByIndex, _LOBase_QueryFieldsGetNames
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
; Parameters ....: $oQuery              - A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning array of Query names. @Extended will be set to the number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oQuery not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Array of Query names.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The name returned will be the Alias of the field, if there is one.
; Related .......: _LOBase_QueryFieldGetObjByName, _LOBase_QueryFieldsGetCount
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
; Parameters ....: $oConnection         - A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $iQuery              - The Index value of the Query to retrieve. 0 Based.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested Query's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oConnection not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oConnection not a Connection Object.
;                  @Error: 1, @Extended: 3 = $iQuery not an Integer, less than 0 or greater than number of Queries contained in the Database.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Connection called in $oConnection is closed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Queries Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Query Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueriesGetCount, _LOBase_QueryGetObjByName, _LOBase_QueryDocOpenByObject
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryGetObjByIndex(ByRef $oConnection, $iQuery)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQuery, $oQueries

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oQueries = $oConnection.Queries()
	If Not IsObj($oQueries) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iQuery, 0, $oQueries.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oQuery = $oQueries.getByIndex($iQuery)
	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oQuery)
EndFunc   ;==>_LOBase_QueryGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QueryGetObjByName
; Description ...: Retrieve a Query's Object by name.
; Syntax ........: _LOBase_QueryGetObjByName(ByRef $oConnection, $sName)
; Parameters ....: $oConnection         - A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sName               - The Query's name to retrieve the Object for.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested Query's Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oConnection not an Object.
;                  @Error: 1, @Extended: 2 = Object called in $oConnection not a Connection Object.
;                  @Error: 1, @Extended: 3 = $sName not a String.
;                  @Error: 1, @Extended: 4 = Query with name called in $sName not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Connection called in $oConnection is closed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Queries Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Query Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_QueryGetObjByIndex, _LOBase_QueriesGetNames, _LOBase_QueryDocOpenByObject
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryGetObjByName(ByRef $oConnection, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oQuery, $oQueries

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
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
; Parameters ....: $oQuery              - A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
;                  $sName               - [optional] Default is Null. The new name to set the Query to. See Remarks.
; Return values .: Success: 1 or String
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. $sName called with Null, returning current Query Name.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oQuery not an Object.
;                  @Error: 1, @Extended: 2 = $sName not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Query's name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function does not check if the new name already exists in Tables or Queries.
;                  According to LibreOffice SDK API IDL XRename Interface, It would seem some Database types don't support the renaming of Queries.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOBase_QueryExists, _LOBase_QueriesGetNames, _LOBase_QueryGetObjByIndex, _LOBase_QueryGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QueryName(ByRef $oQuery, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sCurName
	Local $iError = 0

	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sName) Then
		$sCurName = $oQuery.Name()
		If Not IsString($sCurName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurName)
	EndIf

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oQuery.rename($sName)
	$iError = ($oQuery.Name() = $sName) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_QueryName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_QuerySQLCommand
; Description ...: Set or Retrieve the Query's SQL command.
; Syntax ........: _LOBase_QuerySQLCommand(ByRef $oQuery[, $sSQL_Command = Null])
; Parameters ....: $oQuery              - A Query object returned by a previous _LOBase_QueryGetObjByName, _LOBase_QueryGetObjByIndex, _LOBase_QueryAddByName, or _LOBase_QueryAddBySQL function.
;                  $sSQL_Command        - [optional] Default is Null. The SQL command to set for the Query.
; Return values .: Success: 1 or String
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. $sSQL_Command called with Null, returning current Query SQL Command.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oQuery not an Object.
;                  @Error: 1, @Extended: 2 = $sSQL_Command not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current SQL command.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sSQL_Command
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_QuerySQLCommand(ByRef $oQuery, $sSQL_Command = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sCurSQL
	Local $iError = 0

	If Not IsObj($oQuery) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sSQL_Command) Then
		$sCurSQL = $oQuery.Command()
		If Not IsString($sCurSQL) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurSQL)
	EndIf

	If Not IsString($sSQL_Command) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oQuery.Command = $sSQL_Command
	$iError = ($oQuery.Command() = $sSQL_Command) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_QuerySQLCommand
