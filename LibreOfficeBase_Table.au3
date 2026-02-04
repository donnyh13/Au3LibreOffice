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
; Description ...: Provides basic functionality through AutoIt for Adding, Deleting, and modifying, etc. L.O. Base Tables.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_TableAdd
; _LOBase_TableColAdd
; _LOBase_TableColDefinition
; _LOBase_TableColDelete
; _LOBase_TableColGetObjByIndex
; _LOBase_TableColGetObjByName
; _LOBase_TableColProperties
; _LOBase_TableColsGetCount
; _LOBase_TableColsGetNames
; _LOBase_TableDelete
; _LOBase_TableExists
; _LOBase_TableGetObjByIndex
; _LOBase_TableGetObjByName
; _LOBase_TableIndexAdd
; _LOBase_TableIndexDelete
; _LOBase_TableIndexesGetCount
; _LOBase_TableIndexesGetNames
; _LOBase_TableIndexModify
; _LOBase_TableName
; _LOBase_TablePrimaryKey
; _LOBase_TablesGetCount
; _LOBase_TablesGetNames
; _LOBase_TableUIClose
; _LOBase_TableUIConnect
; _LOBase_TableUIGetRowSet
; _LOBase_TableUIOpenByName
; _LOBase_TableUIOpenByObject
; _LOBase_TableUIVisible
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableAdd
; Description ...: Add a Table to a Database.
; Syntax ........: _LOBase_TableAdd(ByRef $oConnection, $sName, $sColName[, $iColType = $LOB_DATA_TYPE_VARCHAR[, $sColTypeName = ""[, $sColDesc = ""]]])
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sName               - a string value. The Unique name of the table to create.
;                  $sColName            - a string value. The Name for the first column.
;                  $iColType            - [optional] an integer value (-16-2014). Default is $LOB_DATA_TYPE_VARCHAR. The new Column's data type. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  $sColTypeName        - [optional] a string value. Default is "". If the column type is a user-defined type, then a fully-qualified type name will be entered here.
;                  $sColDesc            - [optional] a string value. Default is "". The description text of the new column.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = $sColName not a String.
;                  @Error 1 @Extended 5 Return 0 = $iColType not an Integer, less than -16 or greater than 2014
;                  @Error 1 @Extended 6 Return 0 = $sColTypeName not a String.
;                  @Error 1 @Extended 7 Return 0 = $sColDesc not a String.
;                  @Error 1 @Extended 8 Return 0 = Table name called in $sName already used as a Table name.
;                  @Error 1 @Extended 9 Return 0 = Table name called in $sName already used as a Query name.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Table Descriptor.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a Column Descriptor.
;                  @Error 2 @Extended 3 Return 0 = Failed to create a Key Descriptor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve appropriate Type Name.
;                  @Error 3 @Extended 2 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Tables Object.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Columns Object.
;                  @Error 3 @Extended 5 Return 0 = Failed to insert new Table.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve New Table's Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning new Table's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The first column created is set as the primary key.
;                  It is the user's responsibility to determine which Data types are valid to be used.
; Related .......: _LOBase_TableDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableAdd(ByRef $oConnection, $sName, $sColName, $iColType = $LOB_DATA_TYPE_VARCHAR, $sColTypeName = "", $sColDesc = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTable, $oTables, $oTableDesc, $oColumns, $oColumn, $oKeyDesc
	Local Const $__LOB_KEY_TYPE_PRIMARY = 1

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sColName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iColType, $LOB_DATA_TYPE_LONGNVARCHAR, $LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsString($sColTypeName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsString($sColDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If ($sColTypeName = "") Then $sColTypeName = __LOBase_ColTypeName($iColType)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oTables = $oConnection.getTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If $oTables.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If $oConnection.Queries.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

	$oTableDesc = $oTables.createDataDescriptor()
	If Not IsObj($oTableDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTableDesc.Name = $sName

	$oColumns = $oTableDesc.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$oColumn = $oColumns.createDataDescriptor()
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	With $oColumn
		.Name = $sColName
		.Type = $iColType
		.TypeName = $sColTypeName
		.HelpText = $sColDesc

	EndWith

	Switch $iColType
		Case $LOB_DATA_TYPE_BOOLEAN
			$oColumn.Precision = 1

		Case $LOB_DATA_TYPE_TINYINT
			$oColumn.Precision = 3

		Case $LOB_DATA_TYPE_SMALLINT, $LOB_DATA_TYPE_FLOAT
			$oColumn.Precision = 5

		Case $LOB_DATA_TYPE_INTEGER
			$oColumn.Precision = 10

		Case $LOB_DATA_TYPE_REAL, $LOB_DATA_TYPE_DOUBLE
			$oColumn.Precision = 17

		Case $LOB_DATA_TYPE_BIGINT
			$oColumn.Precision = 19

		Case $LOB_DATA_TYPE_LONGVARBINARY, $LOB_DATA_TYPE_VARBINARY, $LOB_DATA_TYPE_LONGNVARCHAR, $LOB_DATA_TYPE_BINARY, $LOB_DATA_TYPE_CHAR, _
				$LOB_DATA_TYPE_VARCHAR, $LOB_DATA_TYPE_NVARCHAR, $LOB_DATA_TYPE_OTHER
			$oColumn.Precision = 2147483647

		Case $LOB_DATA_TYPE_NUMERIC, $LOB_DATA_TYPE_DECIMAL
			$oColumn.Precision = 646456993

;~ 		Case $LOB_DATA_TYPE_DATE, $LOB_DATA_TYPE_TIME, $LOB_DATA_TYPE_TIMESTAMP; No value needed.
	EndSwitch

	$oColumns.appendByDescriptor($oColumn)

	$oTables.appendByDescriptor($oTableDesc)

	If Not $oTables.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

	$oTable = $oTables.getByName($sName)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

	; Set Primary Key.
	$oKeyDesc = $oTable.Keys().createDataDescriptor()
	If Not IsObj($oKeyDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	$oKeyDesc.Columns().appendByDescriptor($oColumn)
	$oKeyDesc.Type = $__LOB_KEY_TYPE_PRIMARY

	$oTable.Keys().appendByDescriptor($oKeyDesc)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTable)
EndFunc   ;==>_LOBase_TableAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableColAdd
; Description ...: Add a new Column to a Table.
; Syntax ........: _LOBase_TableColAdd(ByRef $oTable, $sName, $iType[, $sTypeName = ""[, $sDescription = ""]])
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
;                  $sName               - a string value. The unique Column Name.
;                  $iType               - an integer value (-16-2014). The Column Type. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  $sTypeName           - [optional] a string value. Default is "". If the column type is a user-defined type, then a fully-qualified type name will be entered here.
;                  $sDescription        - [optional] a string value. Default is "". The description text of the new column.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = $iType not an Integer, less than -16 or greater than 2014. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $sTypeName not a String.
;                  @Error 1 @Extended 5 Return 0 = $sDescription not a String.
;                  @Error 1 @Extended 6 Return 0 = Column with the same name as called in $sName already exists.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Column descriptor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve appropriate Type Name.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Columns Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to insert new Column.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve new Column's Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully inserted the new column, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableColDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableColAdd(ByRef $oTable, $sName, $iType, $sTypeName = "", $sDescription = "")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns, $oColumn

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iType, $LOB_DATA_TYPE_LONGNVARCHAR, $LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sTypeName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sDescription) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If ($sTypeName = "") Then $sTypeName = __LOBase_ColTypeName($iType)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oColumns = $oTable.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If $oColumns.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oColumn = $oColumns.createDataDescriptor()
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oColumn.Name = $sName
	$oColumn.Type = $iType
	$oColumn.TypeName = $sTypeName
	$oColumn.HelpText = $sDescription

	Switch $iType
		Case $LOB_DATA_TYPE_BOOLEAN
			$oColumn.Precision = 1

		Case $LOB_DATA_TYPE_TINYINT
			$oColumn.Precision = 3

		Case $LOB_DATA_TYPE_SMALLINT, $LOB_DATA_TYPE_FLOAT
			$oColumn.Precision = 5

		Case $LOB_DATA_TYPE_INTEGER
			$oColumn.Precision = 10

		Case $LOB_DATA_TYPE_REAL, $LOB_DATA_TYPE_DOUBLE
			$oColumn.Precision = 17

		Case $LOB_DATA_TYPE_BIGINT
			$oColumn.Precision = 19

		Case $LOB_DATA_TYPE_LONGVARBINARY, $LOB_DATA_TYPE_VARBINARY, $LOB_DATA_TYPE_LONGNVARCHAR, $LOB_DATA_TYPE_BINARY, $LOB_DATA_TYPE_CHAR, _
				$LOB_DATA_TYPE_VARCHAR, $LOB_DATA_TYPE_NVARCHAR, $LOB_DATA_TYPE_OTHER
			$oColumn.Precision = 2147483647

		Case $LOB_DATA_TYPE_NUMERIC, $LOB_DATA_TYPE_DECIMAL
			$oColumn.Precision = 646456993
	EndSwitch

	$oColumns.appendByDescriptor($oColumn)

	If Not $oColumns.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oColumn = $oColumns.getByName($sName)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOBase_TableColAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableColDefinition
; Description ...: Set or Retrieve column type settings.
; Syntax ........: _LOBase_TableColDefinition(ByRef $oTable, ByRef $oColumn[, $sName = Null[, $iType = Null[, $sTypeName = Null[, $sDescription = Null]]]])
; Parameters ....: $oTable              - [in/out] an object.A Table object returned by a previous _LOBase_TableGetObjByIndex or _LOBase_TableGetObjByName function.
;                  $oColumn             - [in/out] an object. A Column object returned by a previous _LOBase_TableColGetObjByIndex or _LOBase_TableColGetObjByName function.
;                  $sName               - [optional] a string value. Default is Null. The Column Name.
;                  $iType               - [optional] an integer value (-16-2014). Default is Null. The Column Type. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  $sTypeName           - [optional] a string value. Default is Null. If the column type is a user-defined type, then a fully-qualified type name will be entered here.
;                  $sDescription        - [optional] a string value. Default is Null. The description text of the column.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oColumn not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = Column with the same name as called in $sName already exists.
;                  @Error 1 @Extended 5 Return 0 = $iType not an Integer, less than -16 or greater than 2014. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $sTypeName not a String.
;                  @Error 1 @Extended 7 Return 0 = $sDescription not a String.
;                  @Error 1 @Extended 8 Return 0 = Column called in $oColumn not a Table Column and does not support a description. See Remarks.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Column descriptor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Old Column name.
;                  @Error 3 @Extended 2 Return 0 = Failed to transfer old column properties to new column Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Appropriate Type name.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve new column Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  |                               2 = Error setting $iType
;                  |                               4 = Error setting $sTypeName
;                  |                               8 = Error setting $sDescription
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 3 or 4 Element Array with values in order of function parameters. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Column Objects retrieved for primary keys do not support a Description text, thus if a Primary Key Column is called in $oColumn, that parameter will be omitted from the returned array when retrieving the settings.
; Related .......: _LOBase_TableColProperties
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableColDefinition(ByRef $oTable, ByRef $oColumn, $sName = Null, $iType = Null, $sTypeName = Null, $sDescription = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNewCol
	Local $sOldName
	Local $iError = 0
	Local $asSettings[3]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($sName, $iType, $sTypeName, $sDescription) Then
		If $oColumn.supportsService("com.sun.star.sdbcx.KeyColumn") Then ; Key Column
			__LO_ArrayFill($asSettings, $oColumn.Name(), $oColumn.Type(), $oColumn.TypeName())

		Else
			__LO_ArrayFill($asSettings, $oColumn.Name(), $oColumn.Type(), $oColumn.TypeName(), $oColumn.HelpText())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $asSettings)
	EndIf

	$sOldName = $oColumn.Name()
	If Not IsString($sOldName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If $oTable.Columns.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oColumn.setName($sName)
		$iError = ($oColumn.Name() = $sName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If Not __LO_VarsAreNull($iType, $sTypeName, $sDescription) Then
		$oNewCol = $oTable.Columns.createDataDescriptor()
		If Not IsObj($oNewCol) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		__LOBase_ColTransferProps($oNewCol, $oColumn)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		If ($iType <> Null) Then
			If Not __LO_IntIsBetween($iType, $LOB_DATA_TYPE_LONGNVARCHAR, $LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$oNewCol.Type = $iType
			If ($sTypeName = Null) Then $sTypeName = __LOBase_ColTypeName($iType)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		EndIf

		If ($sTypeName <> Null) Then
			If Not IsString($sTypeName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$oNewCol.TypeName = $sTypeName
		EndIf

		$oTable.alterColumnByName($sOldName, $oNewCol)

		$oNewCol = $oTable.Columns.getByName($sName)
		If Not IsObj($oNewCol) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		$oColumn = $oNewCol

		$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oColumn.Type() = $iType) ? ($iError) : (BitOR($iError, 2)))
		$iError = (__LO_VarsAreNull($sTypeName)) ? ($iError) : (($oColumn.TypeName() = $sTypeName) ? ($iError) : (BitOR($iError, 4)))
	EndIf

	If ($sDescription <> Null) Then
		If Not IsString($sDescription) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		If Not $oColumn.supportsService("com.sun.star.sdbcx.Column") Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0) ; Normal Column

		$oNewCol.HelpText = $sDescription
		$iError = ($oColumn.HelpText() = $sDescription) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_TableColDefinition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableColDelete
; Description ...: Delete a Column from a Table.
; Syntax ........: _LOBase_TableColDelete(ByRef $oTable, ByRef $oColumn)
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
;                  $oColumn             - [in/out] an object. A Column object returned by a previous _LOBase_TableColGetObjByIndex or _LOBase_TableColGetObjByName function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oColumn not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Column Name.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Columns Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to delete Column.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Column was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableColAdd
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableColDelete(ByRef $oTable, ByRef $oColumn)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns, $sName

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sName = $oColumn.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oColumns = $oTable.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oColumns.dropByName($sName)

	If $oColumns.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_TableColDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableColGetObjByIndex
; Description ...: Retrieve a Table Column's Object by Index.
; Syntax ........: _LOBase_TableColGetObjByIndex(ByRef $oTable, $iIndex)
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
;                  $iIndex              - an integer value. The Index of the Column to retrieve the Column for. 0 Based.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iIndex not an Integer, less than 0 or greater than number of Columns.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Column Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Column's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableColsGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableColGetObjByIndex(ByRef $oTable, $iIndex)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns, $oColumn

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oTable.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iIndex, 0, $oColumns.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumn = $oColumns.getByIndex($iIndex)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOBase_TableColGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableColGetObjByName
; Description ...: Retrieve a Table Column's Object by name.
; Syntax ........: _LOBase_TableColGetObjByName(ByRef $oTable, $sName)
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
;                  $sName               - a string value. The name of the Column to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Table does not contain a column with a name as called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Column Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Column's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: I found one place online stating more than one column can have the same name, and thus, this could be an unreliable method of obtaining the Column's object, in the case that there are two columns identically named. However I have not been able to reproduce this behavior.
; Related .......: _LOBase_TableColsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableColGetObjByName(ByRef $oTable, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns, $oColumn

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oColumns = $oTable.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not $oColumns.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oColumn = $oColumns.getByName($sName)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oColumn)
EndFunc   ;==>_LOBase_TableColGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableColProperties
; Description ...: Set or Retrieve Column properties.
; Syntax ........: _LOBase_TableColProperties(ByRef $oConnection, ByRef $oTable, ByRef $oColumn[, $iLength = Null[, $sDefaultVal = Null[, $bRequired = Null[, $iDecimalPlace = Null[, $bAutoValue = Null[, $iFormat = Null[, $iAlign = Null]]]]]]])
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
;                  $oColumn             - [in/out] an object. A Column object returned by a previous _LOBase_TableColGetObjByIndex or _LOBase_TableColGetObjByName function.
;                  $iLength             - [optional] an integer value. Default is Null. The maximum number of characters allowed to be entered.
;                  $sDefaultVal         - [optional] a string value. Default is Null. The Default value of the column. See remarks.
;                  $bRequired           - [optional] a boolean value. Default is Null. If True, the column cannot be empty.
;                  $iDecimalPlace       - [optional] an integer value (0-32767). Default is Null. The Decimal place for numerical values.
;                  $bAutoValue          - [optional] a boolean value. Default is Null. If True, The column's value is auto-generated.
;                  $iFormat             - [optional] an integer value. Default is Null. The Number Format Key to display the content in, retrieved from a previous _LOBase_FormatKeysGetList call, or created by _LOBase_FormatKeyCreate function.
;                  $iAlign              - [optional] an integer value (0-2). Default is Null. The alignment of the column's text. See Constants, $LOB_COL_TXT_ALIGN_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 4 Return 0 = $oColumn not an Object.
;                  @Error 1 @Extended 5 Return 0 = $iLength not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $sDefaultVal not a String.
;                  @Error 1 @Extended 7 Return 0 = $bRequired not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $iDecimalPlace not an Integer, less than 0 or greater than 32,767.
;                  @Error 1 @Extended 9 Return 0 = $bAutoValue not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $iFormat not an Integer.
;                  @Error 1 @Extended 11 Return 0 = Format key called in $iFormat not found.
;                  @Error 1 @Extended 12 Return 0 = $iAlign not an Integer, less than 0 or greater than 2. See Constants, $LOB_COL_TXT_ALIGN_* as defined in LibreOfficeBase_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a Column descriptor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to transfer old column properties to new column Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve new column object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLength
;                  |                               2 = Error setting $sDefaultVal
;                  |                               4 = Error setting $bRequired
;                  |                               8 = Error setting $iDecimalPlace
;                  |                               16 = Error setting $bAutoValue
;                  |                               32 = Error setting $iFormat
;                  |                               64 = Error setting $iAlign
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: For $sDefaultVal, enter any numerical values as strings.
;                  Not all column types support all of these settings. It is the user's responsibility to know which are valid or not.
;                  There seems to be Constant value for Default, Justified and Filled settings for $iAlign.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_TableColDefinition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableColProperties(ByRef $oConnection, ByRef $oTable, ByRef $oColumn, $iLength = Null, $sDefaultVal = Null, $bRequired = Null, $iDecimalPlace = Null, $bAutoValue = Null, $iFormat = Null, $iAlign = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $bAlter = False
	Local $oNewCol
	Local $asSettings[7]
	Local Const $__LOB_IS_REQUIRED_YES = 0, $__LOB_IS_REQUIRED_NO = 1

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If __LO_VarsAreNull($iLength, $sDefaultVal, $bRequired, $iDecimalPlace, $bAutoValue, $iFormat, $iAlign) Then
		__LO_ArrayFill($asSettings, $oColumn.Precision(), $oColumn.ControlDefault(), _
				($oColumn.IsNullable() = $__LOB_IS_REQUIRED_YES) ? (True) : (False), _
				$oColumn.Scale(), $oColumn.IsAutoIncrement(), $oColumn.FormatKey(), $oColumn.Align())

		Return SetError($__LO_STATUS_SUCCESS, 1, $asSettings)
	EndIf

	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oNewCol = $oTable.Columns.createDataDescriptor()
	If Not IsObj($oNewCol) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	__LOBase_ColTransferProps($oNewCol, $oColumn)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If ($iLength <> Null) Then
		If Not IsInt($iLength) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oNewCol.Precision = $iLength
		$bAlter = True
	EndIf

	If ($sDefaultVal <> Null) Then
		If Not IsString($sDefaultVal) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oColumn.ControlDefault = $sDefaultVal
		$iError = ($oColumn.ControlDefault() = $sDefaultVal) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bRequired <> Null) Then
		If Not IsBool($bRequired) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oNewCol.IsNullable = ($bRequired) ? ($__LOB_IS_REQUIRED_YES) : ($__LOB_IS_REQUIRED_NO)
		$bAlter = True
	EndIf

	If ($iDecimalPlace <> Null) Then
		If Not __LO_IntIsBetween($iDecimalPlace, 0, 32767) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oNewCol.Scale = $iDecimalPlace
		$bAlter = True
	EndIf

	If ($bAutoValue <> Null) Then
		If Not IsBool($bAutoValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oNewCol.IsAutoIncrement = $bAutoValue
		$bAlter = True
	EndIf

	If $bAlter Then
		$oTable.alterColumnByName($oColumn.Name(), $oNewCol)

		$oNewCol = $oTable.Columns.getByName($oColumn.Name())
		If Not IsObj($oNewCol) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oColumn = $oNewCol
	EndIf

	; Have to set these after altering Column, else they are erased.
	If ($iFormat <> Null) Then
		If Not IsInt($iFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		If Not _LOBase_FormatKeyExists($oConnection, $iFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$oColumn.FormatKey = $iFormat
		$iError = ($oColumn.FormatKey() = $iFormat) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iAlign <> Null) Then
		If Not __LO_IntIsBetween($iAlign, $LOB_COL_TXT_ALIGN_LEFT, $LOB_COL_TXT_ALIGN_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oColumn.Align = $iAlign
		$iError = ($oColumn.Align() = $iAlign) ? ($iError) : (BitOR($iError, 64))
	EndIf

	ConsoleWrite($oNewCol.Align() & @CRLF)

	$iError = (__LO_VarsAreNull($iLength)) ? ($iError) : (($oColumn.Precision() = $iLength) ? ($iError) : (BitOR($iError, 1)))
	$iError = (__LO_VarsAreNull($bRequired)) ? ($iError) : (($oColumn.IsNullable() = (($bRequired) ? ($__LOB_IS_REQUIRED_YES) : ($__LOB_IS_REQUIRED_NO))) ? ($iError) : (BitOR($iError, 4)))
	$iError = (__LO_VarsAreNull($iDecimalPlace)) ? ($iError) : (($oColumn.Scale() = $iDecimalPlace) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($bAutoValue)) ? ($iError) : (($oColumn.IsAutoIncrement() = $bAutoValue) ? ($iError) : (BitOR($iError, 16)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_TableColProperties

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableColsGetCount
; Description ...: Retrieve a count of the number of columns contained in a Table.
; Syntax ........: _LOBase_TableColsGetCount(ByRef $oTable)
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve count of columns contained in the Table.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Columns contained in the Table.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableColGetObjByIndex
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableColsGetCount(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns
	Local $iCount

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oTable.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iCount = $oColumns.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOBase_TableColsGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableColsGetNames
; Description ...: Retrieve an array of Column names contained in a Table.
; Syntax ........: _LOBase_TableColsGetNames(ByRef $oTable)
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Columns Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Array of Column names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Column names. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableColGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableColsGetNames(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumns
	Local $asNames[0]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oColumns = $oTable.getColumns()
	If Not IsObj($oColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asNames = $oColumns.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOBase_TableColsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableDelete
; Description ...: Delete a Table from a Database.
; Syntax ........: _LOBase_TableDelete(ByRef $oConnection, ByRef $oTable)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Tables Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Table name.
;                  @Error 3 @Extended 4 Return 0 = Failed to delete Table.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Table was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableAdd
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableDelete(ByRef $oConnection, ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTables
	Local $sName

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oTables = $oConnection.getTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$sName = $oTable.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oTables.dropByName($sName)

	If $oTables.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_TableDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableExists
; Description ...: Check whether a Database contains a Table by name.
; Syntax ........: _LOBase_TableExists(ByRef $oConnection, $sName)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sName               - a string value. The name of the Table to look for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Tables Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to query Tables for Table name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning a Boolean value indicating if the Database contains a Table by the called name (True) or not.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableExists(ByRef $oConnection, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTables
	Local $bReturn

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oTables = $oConnection.getTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$bReturn = $oTables.hasByName($sName)
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_TableExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableGetObjByIndex
; Description ...: Retrieve a Table's Object by Index.
; Syntax ........: _LOBase_TableGetObjByIndex(ByRef $oConnection, $iTable)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $iTable              - an integer value. The Index value of the Table to retrieve. 0 Based.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $iTable not an Integer, less than 0 or greater than number of Tables contained in the Database.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Tables Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Table Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Table's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TablesGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableGetObjByIndex(ByRef $oConnection, $iTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTable, $oTables

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oTables = $oConnection.getTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iTable, 0, $oTables.Count() - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTable = $oTables.getByIndex($iTable)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTable)
EndFunc   ;==>_LOBase_TableGetObjByIndex

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableGetObjByName
; Description ...: Retrieve a Table's Object by name.
; Syntax ........: _LOBase_TableGetObjByName(ByRef $oConnection, $sName)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sName               - a string value. The Table's name to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $sName not a String.
;                  @Error 1 @Extended 4 Return 0 = Table with name called in $sName not found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Tables Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Table Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Table's Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TablesGetNames, _LOBase_TableExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableGetObjByName(ByRef $oConnection, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTable, $oTables

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oTables = $oConnection.getTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oTables.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oTable = $oTables.getByName($sName)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTable)
EndFunc   ;==>_LOBase_TableGetObjByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableIndexAdd
; Description ...: Add an Index to a Table.
; Syntax ........: _LOBase_TableIndexAdd(ByRef $oTable, $sName, $avColumns[, $bIsUnique = False])
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex or _LOBase_TableGetObjByName function.
;                  $sName               - a string value. The name of the new Index.
;                  $avColumns           - an array of variants. A 2 column array of Column names and accompanying Boolean values. See Remarks.
;                  $bIsUnique           - [optional] a boolean value. Default is False. If True the Indexed Column(s) can contain only unique entries.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Table does not have an Index with the name called in $sName.
;                  @Error 1 @Extended 4 Return 0 = $avColumns is not an Array, or has 0 Elements, or does not contain 2 columns.
;                  @Error 1 @Extended 5 Return 0 = $bIsUnique not a Boolean.
;                  @Error 1 @Extended 6 Return ? = Column 1 (0th Column) of $avColumns contains a non-string. Returning problem Element number.
;                  @Error 1 @Extended 7 Return ? = Column name called in Column 1 (0th Column) of $avColumns does not exist in Table. Returning problem Element number.
;                  @Error 1 @Extended 8 Return ? = Column 2 (1) of $avColumns contains a non-Boolean value. Returning problem Element number.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create an Index Descriptor.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a Column Descriptor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to append the new Index.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. New Index was successfully added to the Table.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array called in $avColumns needs to be a 2 Column array, the Column name must be placed in the first (0th) column, and a Boolean value indicating whether the Column should should be sorted Ascending (True) or Descending (False) be found in the second (1st) column.
;                  An example of creating an Array for $avColumns would be: Local $avColumns[1][2] = [["ColumnName", [True]]. This would sort the Column named "ColumnName" in Ascending order.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableIndexAdd(ByRef $oTable, $sName, $avColumns, $bIsUnique = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oIndexDesc, $oColumnDesc
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oTable.Indexes.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsArray($avColumns) Or (UBound($avColumns) < 1) Or (UBound($avColumns, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bIsUnique) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oIndexDesc = $oTable.Indexes.createDataDescriptor()
	If Not IsObj($oIndexDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oIndexDesc.Name = $sName
	$oIndexDesc.IsUnique = $bIsUnique

	For $i = 0 To UBound($avColumns) - 1
		If Not IsString($avColumns[$i][0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)
		If Not $oTable.Columns.hasByName($avColumns[$i][0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, $i)
		If Not IsBool($avColumns[$i][1]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, $i)

		$oColumnDesc = $oIndexDesc.Columns.createDataDescriptor()
		If Not IsObj($oColumnDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oColumnDesc.setName($avColumns[$i][0])
		$oColumnDesc.IsAscending = $avColumns[$i][1]
		$oIndexDesc.Columns.appendByDescriptor($oColumnDesc)
		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	$oTable.Indexes.appendByDescriptor($oIndexDesc)
	If Not $oTable.Indexes.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_TableIndexAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableIndexDelete
; Description ...: Delete a Table Index by name.
; Syntax ........: _LOBase_TableIndexDelete(ByRef $oTable, $sName)
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex or _LOBase_TableGetObjByName function.
;                  $sName               - a string value. The Index name to delete.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Table called in $oTable does not contain an Index with the name called in $sName.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to delete the Index.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully deleted the Index.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableIndexDelete(ByRef $oTable, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oTable.Indexes.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTable.Indexes.dropByName($sName)
	If $oTable.Indexes.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_TableIndexDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableIndexesGetCount
; Description ...: Retrieve a count of Indexes for a Table.
; Syntax ........: _LOBase_TableIndexesGetCount(ByRef $oTable)
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex or _LOBase_TableGetObjByName function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Index Count.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Indexes contained in the Table.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableIndexesGetCount(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCount = $oTable.Indexes.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOBase_TableIndexesGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableIndexesGetNames
; Description ...: Retrieve an array of Table Index names.
; Syntax ........: _LOBase_TableIndexesGetNames(ByRef $oTable)
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex or _LOBase_TableGetObjByName function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve array of Index names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning an Array of Index names. @Extended is set to the number of Elements contained in the Array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableIndexesGetNames(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asIndex[0]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asIndex = $oTable.Indexes.ElementNames()
	If Not IsArray($asIndex) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asIndex), $asIndex)
EndFunc   ;==>_LOBase_TableIndexesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableIndexModify
; Description ...: Modify the columns used in an Index.
; Syntax ........: _LOBase_TableIndexModify(ByRef $oTable, $sName[, $avColumns = Null[, $bIsUnique = Null]])
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex or _LOBase_TableGetObjByName function.
;                  $sName               - a string value. The Index name to modify.
;                  $avColumns           - [optional] an array of variants. Default is Null. A 2 column array of Column names and accompanying Boolean values. See Remarks.
;                  $bIsUnique           - [optional] a boolean value. Default is Null. If True the Indexed Column(s) can contain only unique entries.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Table called in $oTable does not contain an Index with the name called in $sName.
;                  @Error 1 @Extended 4 Return 0 = $avColumns is not an Array, or has 0 Elements, or does not contain 2 columns.
;                  @Error 1 @Extended 5 Return ? = Column 1 (0th Column) of $avColumns contains a non-string. Returning problem Element number.
;                  @Error 1 @Extended 6 Return ? = Column name called in Column 1 (0th Column) of $avColumns does not exist in Table. Returning problem Element number.
;                  @Error 1 @Extended 7 Return ? = Column 2 (1) of $avColumns contains a non-Boolean value. Returning problem Element number.
;                  @Error 1 @Extended 8 Return 0 = $bIsUnique not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create an Index Descriptor.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a Column Descriptor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the Index's Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve a Column Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve array of Column names contained in the Index.
;                  @Error 3 @Extended 4 Return 0 =Failed to delete old Index.
;                  @Error 3 @Extended 5 Return 0 = Failed to add modified Index.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve new Index Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColumns
;                  |                               2 = Error setting $bIsUnique
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array called in $avColumns needs to be a 2 Column array, the Column name must be placed in the first (0th) column, and a Boolean value indicating whether the Column should should be sorted Ascending (True) or Descending (False) be found in the second (1st) column.
;                  An example of creating an Array for $avColumns would be: Local $avColumns[1][2] = [["ColumnName", [True]]. This would sort the Column named "ColumnName" in Ascending order.
;                  When retrieving the current settings, the returned array will be as described above for the $avColumns value.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  The error checking for newly set Columns or Ascending/Descending values doesn't check the content of the Index's columns vs those called in $avColumns, only the number of Columns.
;                  According to LibreOffice SDK API, some databases ignore the Ascending/Descending settings. In my limited testing, embedded HSQLDB seems to always be set to Ascending.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableIndexModify(ByRef $oTable, $sName, $avColumns = Null, $bIsUnique = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oColumn, $oColumnDesc, $oIndex, $oIndexDesc
	Local $avSettings[2], $avCurrentColumns[0][2], $asIndexColumns[0]
	Local $bDelete = True
	Local Const $__STR_STRIPLEADING = 1, $__STR_STRIPTRAILING = 2, $__UBOUND_COLUMNS = 2
	Local $iError = 0

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oTable.Indexes.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oIndex = $oTable.Indexes.getByName($sName)
	If Not IsObj($oIndex) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $avCurrentColumns[$oIndex.Columns.Count()][2]

	For $i = 0 To $oIndex.Columns.Count() - 1
		$oColumn = $oIndex.Columns.getByIndex($i)
		If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$avCurrentColumns[$i][0] = $oColumn.Name()
		$avCurrentColumns[$i][1] = $oColumn.IsAscending()
		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	If __LO_VarsAreNull($avColumns, $bIsUnique) Then
		__LO_ArrayFill($avSettings, $avCurrentColumns, $oIndex.IsUnique())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSettings)
	EndIf

	If ($avColumns <> Null) Then
		If Not IsArray($avColumns) Or (UBound($avColumns) < 1) Or (UBound($avColumns, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		For $i = 0 To UBound($avColumns) - 1
			If Not IsString($avColumns[$i][0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)
			If Not $oTable.Columns.hasByName($avColumns[$i][0]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, $i)
			If Not IsBool($avColumns[$i][1]) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, $i)

			If $oIndex.Columns.hasByName($avColumns[$i][0]) Then
				$oColumn = $oIndex.Columns.getByName($avColumns[$i][0])
				If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				If ($oColumn.IsAscending() <> $avColumns[$i][1]) Then
					$oIndex.Columns.dropByName($avColumns[$i][0])

					$oColumnDesc = $oIndex.Columns.createDataDescriptor()
					If Not IsObj($oColumnDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$oColumnDesc.setName($avColumns[$i][0])
					$oColumnDesc.IsAscending = $avColumns[$i][1]
					$oIndex.Columns.appendByDescriptor($oColumnDesc)
				EndIf

			Else
				$oColumnDesc = $oIndex.Columns.createDataDescriptor()
				If Not IsObj($oColumnDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

				$oColumnDesc.setName($avColumns[$i][0])
				$oColumnDesc.IsAscending = $avColumns[$i][1]
				$oIndex.Columns.appendByDescriptor($oColumnDesc)
			EndIf

			Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		$asIndexColumns = $oIndex.Columns.ElementNames()
		If Not IsArray($asIndexColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		For $i = 0 To UBound($asIndexColumns) - 1
			For $k = 0 To UBound($avColumns) - 1
				If (StringStripWS($avColumns[$k][0], ($__STR_STRIPLEADING + $__STR_STRIPTRAILING)) = $asIndexColumns[$i]) Then
					$bDelete = False
					ExitLoop
				EndIf
			Next
			If $bDelete Then $oIndex.Columns.dropByName($asIndexColumns[$i])
			$bDelete = True
			Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
		Next
		$iError = ($oIndex.Columns.Count() = UBound($avColumns)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bIsUnique <> Null) Then
		If Not IsBool($bIsUnique) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		If ($oIndex.IsUnique() <> $bIsUnique) Then
			$oIndexDesc = $oTable.Indexes.createDataDescriptor()
			If Not IsObj($oIndexDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$oIndexDesc.Name = $sName
			$oIndexDesc.IsUnique = $bIsUnique

			For $i = 0 To $oIndex.Columns.Count() - 1
				$oColumn = $oIndex.Columns.getByIndex($i)
				If Not IsObj($oColumn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				$oIndexDesc.Columns.appendByDescriptor($oColumn)
				Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$oTable.Indexes.dropByName($sName)

			If $oTable.Indexes.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			$oTable.Indexes.appendByDescriptor($oIndexDesc)
			If Not $oTable.Indexes.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

			$oIndex = $oTable.Indexes.getByName($sName)
			If Not IsObj($oIndex) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
		EndIf
		$iError = ($oIndex.IsUnique() = $bIsUnique) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_TableIndexModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableName
; Description ...: Set or Retrieve the Table's name.
; Syntax ........: _LOBase_TableName(ByRef $oTable[, $sName = Null])
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
;                  $sName               - [optional] a string value. Default is Null. The new name to set the Table to. See Remarks.
; Return values .: Success: 1 or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sName
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return String = Success. $sName called with Null, returning current Table Name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function does not check if the new name already exists in Tables or Queries.
;                  According to LibreOffice SDK API IDL XRename Interface, It would seem some Database types don't support the renaming of Tables.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOBase_TableExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableName(ByRef $oTable, $sName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sName) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oTable.Name())

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTable.rename($sName)
	If ($oTable.Name() <> $sName) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_TableName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TablePrimaryKey
; Description ...: Set or Retrieve the primary key for a Table.
; Syntax ........: _LOBase_TablePrimaryKey(ByRef $oTable[, $aoPrimary = Null])
; Parameters ....: $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
;                  $aoPrimary           - [optional] an array of objects. Default is Null. An array containing Column Objects (Returned from a previous _LOBase_TableColGetObjByIndex or _LOBase_TableColGetObjByName function).
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $aoPrimary not an Array.
;                  @Error 1 @Extended 3 Return ? = $aoPrimary contains an element that is not a Column Object. Returning problem Element number.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Keys Object
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Primary Key Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Columns Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Array of Column Objects that are currently set as the Primary key.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: There is generally only one Primary key, however it is possible to set more than one Primary key.
;                  When setting only one Column as the Primary key, $aoPrimary must still be an Array.
; Related .......: _LOBase_TableColGetObjByName, _LOBase_TableColGetObjByIndex, _LOBase_TableColAdd
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TablePrimaryKey(ByRef $oTable, $aoPrimary = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oKeys, $oKeysColumns, $oKeyDesc, $oPrimary
	Local $aoPrimaryKeys[0]
	Local Const $__LOB_KEY_TYPE_PRIMARY = 1

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oKeys = $oTable.Keys()
	If Not IsObj($oKeys) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To $oKeys.Count() - 1
		If ($oKeys.getByIndex($i).Type() = $__LOB_KEY_TYPE_PRIMARY) Then
			$oPrimary = $oKeys.getByIndex($i)
			If Not IsObj($oPrimary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			ExitLoop
		EndIf
		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	If __LO_VarsAreNull($aoPrimary) Then
		If IsObj($oPrimary) Then
			$oKeysColumns = $oPrimary.Columns()
			If Not IsObj($oKeysColumns) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			ReDim $aoPrimaryKeys[$oKeysColumns.Count()]

			For $k = 0 To $oKeysColumns.Count() - 1
				$aoPrimaryKeys[$k] = $oKeysColumns.getByIndex($k)
			Next
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, UBound($aoPrimaryKeys), $aoPrimaryKeys)
	EndIf

	If Not IsArray($aoPrimary) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If IsObj($oPrimary) Then
		For $i = 0 To $oPrimary.Columns.Count() - 1
			$oPrimary.Columns.dropByIndex(0)
		Next

		For $k = 0 To UBound($aoPrimary) - 1
			If Not $aoPrimary[$k].supportsService("com.sun.star.sdbcx.Column") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, $i) ; Not a Column Obj

			$oPrimary.Columns().appendByDescriptor($aoPrimary[$k])
		Next

	Else
		$oKeyDesc = $oKeys.createDataDescriptor()
		If Not IsObj($oKeyDesc) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		For $i = 0 To UBound($aoPrimary) - 1
			If Not $aoPrimary[$i].supportsService("com.sun.star.sdbcx.Column") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, $i) ; Not a Column Obj

			$oKeyDesc.Columns().appendByDescriptor($aoPrimary[$i])
		Next

		$oKeyDesc.Type = $__LOB_KEY_TYPE_PRIMARY

		$oTable.Keys().appendByDescriptor($oKeyDesc)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_TablePrimaryKey

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TablesGetCount
; Description ...: Retrieve a count of Tables contained in the Database.
; Syntax ........: _LOBase_TablesGetCount(ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a Connection Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Tables Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve count of Tables.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning count of Tables contained in the Database as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableGetObjByIndex
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TablesGetCount(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTables
	Local $iCount

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oTables = $oConnection.getTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iCount = $oTables.Count()
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCount)
EndFunc   ;==>_LOBase_TablesGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TablesGetNames
; Description ...: Retrieve an Array of Table Names contained in a Database.
; Syntax ........: _LOBase_TablesGetNames(ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a Connection Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Array of Element names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Table names contained in this Database. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableGetObjByName
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TablesGetNames(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asNames[0]

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$asNames = $oConnection.Tables.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOBase_TablesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableUIClose
; Description ...: Close a Table User Interface window.
; Syntax ........: _LOBase_TableUIClose(ByRef $oTableUI[, $bDeliverOwnership = True])
; Parameters ....: $oTableUI            - [in/out] an object. A Table User Interface Object from a previous _LOBase_TableUIOpenByName, _LOBase_TableUIOpenByObject or _LOBase_TableUIConnect function.
;                  $bDeliverOwnership   - [optional] a boolean value. Default is True. If True, deliver ownership of the Table UI Object from the script to LibreOffice, recommended is True.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTableUI not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bDeliverOwnership not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully closed the Table User Interface window.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableUIOpenByObject, _LOBase_TableUIOpenByName, _LOBase_TableUIConnect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableUIClose(ByRef $oTableUI, $bDeliverOwnership = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oTableUI) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bDeliverOwnership) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTableUI.Frame.close($bDeliverOwnership)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_TableUIClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableUIConnect
; Description ...: Connect to an open instance of a Database Table User Interface.
; Syntax ........: _LOBase_TableUIConnect([$bConnectCurrent = True])
; Parameters ....: $bConnectCurrent     - [optional] a boolean value. Default is True. If True, returns the currently active, or last active Document, unless it is not a TableUI Document.
; Return values .: Success: Object or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bConnectCurrent not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating ServiceManager object.
;                  @Error 2 @Extended 2 Return 0 = Error creating Desktop object.
;                  @Error 2 @Extended 3 Return 0 = Error creating enumeration of open documents.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = No open Libre Office documents found.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Row Set Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Table name.
;                  @Error 3 @Extended 4 Return 0 = Current Component not a TableUI Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success, The Object for the current, or last active TableUI document is returned. The Table is open in Viewing/Data entry mode.
;                  @Error 0 @Extended 1 Return Object = Success, The Object for the current, or last active document is returned. The Table is open in Design mode.
;                  @Error 0 @Extended ? Return Array = Success, An Array of all open LibreOffice TableUI documents is returned. See remarks. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Connect all option returns an array with three columns per result. ($aArray[0][3]).
;                  Row 1, Column 0 contains the Object for that document. e.g. $aArray[0][0] = $oDoc
;                  Row 1, Column 1 contains the Document's full title. e.g. $aArray[0][1] = "Table1 - DBaseName" [Viewing mode] OR "DBaseName.odb : Table1" [Design Mode]
;                  Row 1, Column 2 contains a Boolean of whether the TableUI is in Design mode [True] or not.. e.g. $aArray[0][2] = True
;                  Row 2, Column 0 contains the Object for the next document. And so on. e.g. $aArray[1][0] = $oDoc2
; Related .......: _LOBase_TableUIOpenByObject, _LOBase_TableUIOpenByName,
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableUIConnect($bConnectCurrent = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local $aoConnectAll[1][3]
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop, $oRowSet
	Local $sTableName
	Local Const $sTableDesignServ = "com.sun.star.sdb.TableDesign", $sTableViewServ = "com.sun.star.sdb.DataSourceBrowser"

	If Not IsBool($bConnectCurrent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	If Not $oDesktop.getComponents.hasElements() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; no L.O open

	$oEnumDoc = $oDesktop.getComponents.createEnumeration()
	If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	If $bConnectCurrent Then
		$oDoc = $oDesktop.currentComponent()

		If $oDoc.supportsService($sTableDesignServ) Then

			Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc)

		ElseIf $oDoc.supportsService($sTableViewServ) Then
			$oRowSet = $oDoc.FormOperations.Cursor
			If Not IsObj($oRowSet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$sTableName = $oRowSet.Command()
			If Not IsString($sTableName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
			If Not $oRowSet.ActiveConnection.Tables.hasByName($sTableName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; Not a Table UI, but perhaps a Query.

			Return SetError($__LO_STATUS_SUCCESS, 0, $oDoc)

		Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
		EndIf
	EndIf

	; Else Connect All.
	$iCount = 0
	While $oEnumDoc.hasMoreElements()
		$oDoc = $oEnumDoc.nextElement()
		If $oDoc.supportsService($sTableDesignServ) Then
			ReDim $aoConnectAll[$iCount + 1][3]
			$aoConnectAll[$iCount][0] = $oDoc
			$aoConnectAll[$iCount][1] = $oDoc.Title()
			$aoConnectAll[$iCount][2] = True    ; True = In Design mode.
			$iCount += 1

		ElseIf $oDoc.supportsService($sTableViewServ) Then
			$oRowSet = $oDoc.FormOperations.Cursor
			If Not IsObj($oRowSet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$sTableName = $oRowSet.Command()
			If Not IsString($sTableName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			If $oRowSet.ActiveConnection.Tables.hasByName($sTableName) Then
				ReDim $aoConnectAll[$iCount + 1][3]
				$aoConnectAll[$iCount][0] = $oDoc
				$aoConnectAll[$iCount][1] = $oDoc.Title()
				$aoConnectAll[$iCount][2] = False ; False = In Viewing mode.
				$iCount += 1
			EndIf
		EndIf

		Sleep(10)
	WEnd

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoConnectAll)
EndFunc   ;==>_LOBase_TableUIConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableUIGetRowSet
; Description ...: Retrieve a Row Set for a Table opened for Data entry/Viewing. See remarks.
; Syntax ........: _LOBase_TableUIGetRowSet(ByRef $oTableUI)
; Parameters ....: $oTableUI            - [in/out] an object. A Table User Interface Object from a previous _LOBase_TableUIOpenByName, _LOBase_TableUIOpenByObject or _LOBase_TableUIConnect function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTableUI not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oTableUI not Table opened in viewing/data entry mode.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve RowSet Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning Table's RowSet Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving the RowSet for the table allows you to manipulate data contained in the Table using _LOBase_SQLResultRowUpdate, etc. functions.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableUIGetRowSet(ByRef $oTableUI)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oResultSet

	If Not IsObj($oTableUI) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oTableUI.supportsService("com.sun.star.sdb.DataSourceBrowser") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oResultSet = $oTableUI.FormOperations.Cursor()
	If Not IsObj($oResultSet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oResultSet)
EndFunc   ;==>_LOBase_TableUIGetRowSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableUIOpenByName
; Description ...: Open a Table's User Interface window either in design mode or viewing mode.
; Syntax ........: _LOBase_TableUIOpenByName(ByRef $oConnection, $sTable[, $bEdit = False[, $bHidden = False]])
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $sTable              - a string value. The Table's name.
;                  $bEdit               - [optional] a boolean value. Default is False. If True, the Table is opened in editing mode to add or remove columns. If False, the table is opened in data viewing mode, to modify Table Data.
;                  $bHidden             - [optional] a boolean value. Default is False. If True, the UI window will be invisible.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 3 Return 0 = $sTable not a String.
;                  @Error 1 @Extended 4 Return 0 = $bEdit not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = No Table with name called in $sTable found.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Tables Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to create a Connection to Database.
;                  @Error 3 @Extended 4 Return 0 = Failed to open Table UI.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully opened Table's User Interface, returning its object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableUIOpenByObject, _LOBase_TableUIConnect, _LOBase_TableUIClose, _LOBase_TableUIVisible
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableUIOpenByName(ByRef $oConnection, $sTable, $bEdit = False, $bHidden = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTables, $oTableUI
	Local $aArgs[1]

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bEdit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oTables = $oConnection.getTables()
	If Not IsObj($oTables) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oTables.hasByName($sTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If Not $oConnection.Parent.DatabaseDocument.CurrentController.isConnected() Then $oConnection.Parent.DatabaseDocument.CurrentController.connect()
	If Not $oConnection.Parent.DatabaseDocument.CurrentController.isConnected() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$aArgs[0] = __LO_SetPropertyValue("Hidden", $bHidden)

	$oTableUI = $oConnection.Parent.DatabaseDocument.CurrentController.loadComponentWithArguments($LOB_SUB_COMP_TYPE_TABLE, $sTable, $bEdit, $aArgs)
	If Not IsObj($oTableUI) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTableUI)
EndFunc   ;==>_LOBase_TableUIOpenByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableUIOpenByObject
; Description ...: Open a Table's User Interface window either in design mode or viewing mode.
; Syntax ........: _LOBase_TableUIOpenByObject(ByRef $oDoc, ByRef $oConnection, ByRef $oTable[, $bEdit = False[, $bHidden = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
;                  $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $oTable              - [in/out] an object. A Table object returned by a previous _LOBase_TableGetObjByIndex, _LOBase_TableGetObjByName or _LOBase_TableAdd function.
;                  $bEdit               - [optional] a boolean value. Default is False. If True, the Table is opened in editing mode to add or remove columns. If False, the table is opened in data viewing mode, to modify Table Data.
;                  $bHidden             - [optional] a boolean value. Default is False. If True, the UI window will be invisible.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 3 Return 0 = Object called in $oConnection not a Connection Object.
;                  @Error 1 @Extended 4 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 5 Return 0 = $bEdit not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bHidden not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection called in $oConnection is closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Table Name.
;                  @Error 3 @Extended 3 Return 0 = Failed to create a Connection to Database.
;                  @Error 3 @Extended 4 Return 0 = Failed to open Table UI.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully opened Table's User Interface, returning its object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_TableUIOpenByName, _LOBase_TableUIConnect, _LOBase_TableUIClose, _LOBase_TableUIVisible
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableUIOpenByObject(ByRef $oDoc, ByRef $oConnection, ByRef $oTable, $bEdit = False, $bHidden = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTableUI
	Local $sTable
	Local $aArgs[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bEdit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sTable = $oTable.Name()
	If Not IsString($sTable) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If Not $oDoc.CurrentController.isConnected() Then $oDoc.CurrentController.connect()
	If Not $oDoc.CurrentController.isConnected() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$aArgs[0] = __LO_SetPropertyValue("Hidden", $bHidden)

	$oTableUI = $oDoc.CurrentController.loadComponentWithArguments($LOB_SUB_COMP_TYPE_TABLE, $sTable, $bEdit, $aArgs)
	If Not IsObj($oTableUI) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTableUI)
EndFunc   ;==>_LOBase_TableUIOpenByObject

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_TableUIVisible
; Description ...: Set or Retrieve Table UI Visibility.
; Syntax ........: _LOBase_TableUIVisible(ByRef $oTableUI[, $bVisible = Null])
; Parameters ....: $oTableUI            - [in/out] an object. A Table User Interface Object from a previous _LOBase_TableUIOpenByName, _LOBase_TableUIOpenByObject or _LOBase_TableUIConnect function.
;                  $bVisible            - [optional] a boolean value. Default is Null. If True, the Table UI Window is visible.
; Return values .: Success: 1 or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTableUI not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bVisible not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current visibility setting.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bVisible
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. All optional parameters were called with Null, returning current visibility setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_TableUIVisible(ByRef $oTableUI, $bVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oTableUI) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bVisible) Then
		$bVisible = $oTableUI.Frame.ContainerWindow.IsVisible()
		If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $bVisible)
	EndIf

	If Not IsBool($bVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTableUI.Frame.ContainerWindow.Visible = $bVisible
	If Not ($oTableUI.Frame.ContainerWindow.IsVisible() = $bVisible) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_TableUIVisible
