#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Base
#include "LibreOfficeBase_Constants.au3"
#include "LibreOfficeBase_Helper.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Various functions for internal data processing, data retrieval, retrieving and applying settings for LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __LOBase_ColTransferProps
; __LOBase_ColTypeName
; __LOBase_DatabaseMetaGetQuery
; __LOBase_InternalComErrorHandler
; __LOBase_ReportConIdentify
; __LOBase_ReportConSetGetFontDesc
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_ColTransferProps
; Description ...: Transfer column properties from one to another.
; Syntax ........: __LOBase_ColTransferProps(ByRef $oNewCol, ByRef $oOldCol)
; Parameters ....: $oNewCol             - [in/out] an object. A new column Object.
;                  $oOldCol             - [in/out] an object. A Column object returned by a previous _LOBase_TableColGetObjByIndex or _LOBase_TableColGetObjByName function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oNewCol not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oOldCol not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Old Column's Properties.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully transferred Column properties.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_ColTransferProps(ByRef $oNewCol, ByRef $oOldCol)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atProperties[0]

	If Not IsObj($oNewCol) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oOldCol) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$atProperties = $oOldCol.getPropertySetInfo.Properties()
	If Not IsArray($atProperties) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($atProperties) - 1
		If ($oOldCol.getPropertyValue($atProperties[$i].Name()) <> "") Then $oNewCol.setPropertyValue($atProperties[$i].Name(), $oOldCol.getPropertyValue($atProperties[$i].Name()))
		Sleep(($i = $__LOBCONST_SLEEP_DIV) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOBase_ColTransferProps

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_ColTypeName
; Description ...: Obtain an appropriate Type Name for a Column Type.
; Syntax ........: __LOBase_ColTypeName($iType)
; Parameters ....: $iType               - an integer value. The Column Type. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iType not an Integer, less than -16 or greater than 2014. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 2 Return 0 = $iType not one of the pre-defined constants.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the Type name corresponding to the Type Constant.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_ColTypeName($iType)
	Local $sType

	If Not __LO_IntIsBetween($iType, $LOB_DATA_TYPE_LONGNVARCHAR, $LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $iType
		Case $LOB_DATA_TYPE_LONGNVARCHAR
			$sType = "LONGNVARCHAR"

		Case $LOB_DATA_TYPE_NCHAR
			$sType = "NCHAR"

		Case $LOB_DATA_TYPE_NVARCHAR
			$sType = "NVARCHAR"

		Case $LOB_DATA_TYPE_ROWID
			$sType = "ROWID"

		Case $LOB_DATA_TYPE_BIT
			$sType = "BIT"

		Case $LOB_DATA_TYPE_TINYINT
			$sType = "TINYINT"

		Case $LOB_DATA_TYPE_BIGINT
			$sType = "BIGINT"

		Case $LOB_DATA_TYPE_LONGVARBINARY
			$sType = "LONGVARBINARY"

		Case $LOB_DATA_TYPE_VARBINARY
			$sType = "VARBINARY"

		Case $LOB_DATA_TYPE_BINARY
			$sType = "BINARY"

		Case $LOB_DATA_TYPE_LONGVARCHAR
			$sType = "LONGVARCHAR"

		Case $LOB_DATA_TYPE_SQLNULL
			$sType = "SQLNULL"

		Case $LOB_DATA_TYPE_CHAR
			$sType = "CHAR"

		Case $LOB_DATA_TYPE_NUMERIC
			$sType = "NUMERIC"

		Case $LOB_DATA_TYPE_DECIMAL
			$sType = "DECIMAL"

		Case $LOB_DATA_TYPE_INTEGER
			$sType = "INTEGER"

		Case $LOB_DATA_TYPE_SMALLINT
			$sType = "SMALLINT"

		Case $LOB_DATA_TYPE_FLOAT
			$sType = "FLOAT"

		Case $LOB_DATA_TYPE_REAL
			$sType = "REAL"

		Case $LOB_DATA_TYPE_DOUBLE
			$sType = "DOUBLE"

		Case $LOB_DATA_TYPE_VARCHAR
			$sType = "VARCHAR"

		Case $LOB_DATA_TYPE_BOOLEAN
			$sType = "BOOLEAN"

		Case $LOB_DATA_TYPE_DATALINK
			$sType = "DATALINK"

		Case $LOB_DATA_TYPE_DATE
			$sType = "DATE"

		Case $LOB_DATA_TYPE_TIME
			$sType = "TIME"

		Case $LOB_DATA_TYPE_TIMESTAMP
			$sType = "TIMESTAMP"

		Case $LOB_DATA_TYPE_OTHER
			$sType = "OTHER"

		Case $LOB_DATA_TYPE_OBJECT
			$sType = "OBJECT"

		Case $LOB_DATA_TYPE_DISTINCT
			$sType = "DISTINCT"

		Case $LOB_DATA_TYPE_STRUCT
			$sType = "STRUCT"

		Case $LOB_DATA_TYPE_ARRAY
			$sType = "ARRAY"

		Case $LOB_DATA_TYPE_BLOB
			$sType = "BLOB"

		Case $LOB_DATA_TYPE_CLOB
			$sType = "CLOB"

		Case $LOB_DATA_TYPE_REF
			$sType = "REF"

		Case $LOB_DATA_TYPE_SQLXML
			$sType = "SQLXML"

		Case $LOB_DATA_TYPE_NCLOB
			$sType = "NCLOB"

		Case $LOB_DATA_TYPE_REF_CURSOR
			$sType = "REF_CURSOR"

		Case $LOB_DATA_TYPE_TIME_WITH_TIMEZONE
			$sType = "TIME_WITH_TIMEZONE"

		Case $LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE
			$sType = "TIMESTAMP_WITH_TIMEZONE"

		Case Else

			Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, $sType)
EndFunc   ;==>__LOBase_ColTypeName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_DatabaseMetaGetQuery
; Description ...: Return the Query command from a Constant value.
; Syntax ........: __LOBase_DatabaseMetaGetQuery($iQuery)
; Parameters ....: $iQuery              - an integer value (0-148). The Query to retrieve the command for. See Constants, $LOB_DBASE_META_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iQuery not an Integer, less than 0 or greater than number of query commands. See Constants, $LOB_DBASE_META_* as defined in LibreOfficeBase_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the requested Query command.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_DatabaseMetaGetQuery($iQuery)
	Local $asMetaQueries[148]

	If Not __LO_IntIsBetween($iQuery, 0, UBound($asMetaQueries)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$asMetaQueries[$LOB_DBASE_META_ALL_PROCEDURES_ARE_CALLABLE] = ".allProceduresAreCallable"
	$asMetaQueries[$LOB_DBASE_META_ALL_TABLES_ARE_SELECTABLE] = ".allTablesAreSelectable"
	$asMetaQueries[$LOB_DBASE_META_DATA_DEFINITION_CAUSES_TRANSACTION_COMMIT] = ".dataDefinitionCausesTransactionCommit"
	$asMetaQueries[$LOB_DBASE_META_DATA_DEFINITION_IGNORED_IN_TRANSACTIONS] = ".dataDefinitionIgnoredInTransactions"
	$asMetaQueries[$LOB_DBASE_META_DELETES_ARE_DETECTED] = ".deletesAreDetected"
	$asMetaQueries[$LOB_DBASE_META_DOES_MAX_ROW_SIZE_INCLUDE_BLOBS] = ".doesMaxRowSizeIncludeBlobs"
	$asMetaQueries[$LOB_DBASE_META_GET_BEST_ROW_ID] = ".getBestRowIdentifier"
	$asMetaQueries[$LOB_DBASE_META_GET_CATALOG_SEPARATOR] = ".getCatalogSeparator"
	$asMetaQueries[$LOB_DBASE_META_GET_CATALOG_TERM] = ".getCatalogTerm"
	$asMetaQueries[$LOB_DBASE_META_GET_CATALOGS] = ".getCatalogs"
	$asMetaQueries[$LOB_DBASE_META_GET_COLS] = ".getColumns"
	$asMetaQueries[$LOB_DBASE_META_GET_COL_PRIVILEGES] = ".getColumnPrivileges"
	$asMetaQueries[$LOB_DBASE_META_GET_CROSS_REF] = ".getCrossReference"
	$asMetaQueries[$LOB_DBASE_META_GET_DATABASE_PRODUCT_NAME] = ".getDatabaseProductName"
	$asMetaQueries[$LOB_DBASE_META_GET_DATABASE_PRODUCT_VERSION] = ".getDatabaseProductVersion"
	$asMetaQueries[$LOB_DBASE_META_GET_DEFAULT_TRANSACTION_ISOLATION] = ".getDefaultTransactionIsolation"
	$asMetaQueries[$LOB_DBASE_META_GET_DRIVER_MAJOR_VERSION] = ".getDriverMajorVersion"
	$asMetaQueries[$LOB_DBASE_META_GET_DRIVER_MINOR_VERSION] = ".getDriverMinorVersion"
	$asMetaQueries[$LOB_DBASE_META_GET_DRIVER_NAME] = ".getDriverName"
	$asMetaQueries[$LOB_DBASE_META_GET_DRIVER_VERSION] = ".getDriverVersion"
	$asMetaQueries[$LOB_DBASE_META_GET_EXPORTED_KEYS] = ".getExportedKeys"
	$asMetaQueries[$LOB_DBASE_META_GET_EXTRA_NAME_CHARS] = ".getExtraNameCharacters"
	$asMetaQueries[$LOB_DBASE_META_GET_IDENTIFIER_QUOTE_STRING] = ".getIdentifierQuoteString"
	$asMetaQueries[$LOB_DBASE_META_GET_IMPORTED_KEYS] = ".getImportedKeys"
	$asMetaQueries[$LOB_DBASE_META_GET_INDEX_INFO] = ".getIndexInfo"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_BINARY_LITERAL_LEN] = ".getMaxBinaryLiteralLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_CATALOG_NAME_LEN] = ".getMaxCatalogNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_CHAR_LITERAL_LEN] = ".getMaxCharLiteralLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COL_NAME_LEN] = ".getMaxColumnNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COLS_IN_GROUP_BY] = ".getMaxColumnsInGroupBy"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COLS_IN_INDEX] = ".getMaxColumnsInIndex"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COLS_IN_ORDER_BY] = ".getMaxColumnsInOrderBy"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COLS_IN_SEL] = ".getMaxColumnsInSelect"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_COLS_IN_TABLE] = ".getMaxColumnsInTable"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_CONNECTIONS] = ".getMaxConnections"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_CURSOR_NAME_LEN] = ".getMaxCursorNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_INDEX_LEN] = ".getMaxIndexLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_PROCEDURE_NAME_LEN] = ".getMaxProcedureNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_ROW_SIZE] = ".getMaxRowSize"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_SCHEMA_NAME_LEN] = ".getMaxSchemaNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_STATEMENT_LEN] = ".getMaxStatementLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_STATEMENTS] = ".getMaxStatements"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_TABLE_NAME_LEN] = ".getMaxTableNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_TABLES_IN_SEL] = ".getMaxTablesInSelect"
	$asMetaQueries[$LOB_DBASE_META_GET_MAX_USER_NAME_LEN] = ".getMaxUserNameLength"
	$asMetaQueries[$LOB_DBASE_META_GET_NUMERIC_FUNCS] = ".getNumericFunctions"
	$asMetaQueries[$LOB_DBASE_META_GET_PRIMARY_KEY] = ".getPrimaryKeys"
	$asMetaQueries[$LOB_DBASE_META_GET_PROCEDURE_COLS] = ".getProcedureColumns"
	$asMetaQueries[$LOB_DBASE_META_GET_PROCEDURE_TERM] = ".getProcedureTerm"
	$asMetaQueries[$LOB_DBASE_META_GET_PROCEDURES] = ".getProcedures"
	$asMetaQueries[$LOB_DBASE_META_GET_SCHEMA_TERM] = ".getSchemaTerm"
	$asMetaQueries[$LOB_DBASE_META_GET_SCHEMAS] = ".getSchemas"
	$asMetaQueries[$LOB_DBASE_META_GET_SEARCH_STRING_ESCAPE] = ".getSearchStringEscape"
	$asMetaQueries[$LOB_DBASE_META_GET_SQL_KEYWORDS] = ".getSQLKeywords"
	$asMetaQueries[$LOB_DBASE_META_GET_STRING_FUNCS] = ".getStringFunctions"
	$asMetaQueries[$LOB_DBASE_META_GET_SYSTEM_FUNCS] = ".getSystemFunctions"
	$asMetaQueries[$LOB_DBASE_META_GET_TABLE_PRIVILEGES] = ".getTablePrivileges"
	$asMetaQueries[$LOB_DBASE_META_GET_TABLE_TYPES] = ".getTableTypes"
	$asMetaQueries[$LOB_DBASE_META_GET_TABLES] = ".getTables"
	$asMetaQueries[$LOB_DBASE_META_GET_TIME_DATE_FUNCS] = ".getTimeDateFunctions"
	$asMetaQueries[$LOB_DBASE_META_GET_TYPE_INFO] = ".getTypeInfo"
	$asMetaQueries[$LOB_DBASE_META_GET_UDTS] = ".getUDTs"
	$asMetaQueries[$LOB_DBASE_META_GET_URL] = ".getURL"
	$asMetaQueries[$LOB_DBASE_META_GET_USERNAME] = ".getUserName"
	$asMetaQueries[$LOB_DBASE_META_GET_VERSION_COLS] = ".getVersionColumns"
	$asMetaQueries[$LOB_DBASE_META_INSERTS_ARE_DETECTED] = ".insertsAreDetected"
	$asMetaQueries[$LOB_DBASE_META_IS_CATALOG_AT_START] = ".isCatalogAtStart"
	$asMetaQueries[$LOB_DBASE_META_IS_READ_ONLY] = ".isReadOnly"
	$asMetaQueries[$LOB_DBASE_META_NULL_PLUS_NON_NULL_IS_NULL] = ".nullPlusNonNullIsNull"
	$asMetaQueries[$LOB_DBASE_META_NULLS_ARE_SORTED_AT_END] = ".nullsAreSortedAtEnd"
	$asMetaQueries[$LOB_DBASE_META_NULLS_ARE_SORTED_AT_START] = ".nullsAreSortedAtStart"
	$asMetaQueries[$LOB_DBASE_META_NULLS_ARE_SORTED_HIGH] = ".nullsAreSortedHigh"
	$asMetaQueries[$LOB_DBASE_META_NULLS_ARE_SORTED_LOW] = ".nullsAreSortedLow"
	$asMetaQueries[$LOB_DBASE_META_OTHERS_DELETES_ARE_VISIBLE] = ".othersDeletesAreVisible"
	$asMetaQueries[$LOB_DBASE_META_OTHERS_INSERTS_ARE_VISIBLE] = ".othersInsertsAreVisible"
	$asMetaQueries[$LOB_DBASE_META_OTHERS_UPDATES_ARE_VISIBLE] = ".othersUpdatesAreVisible"
	$asMetaQueries[$LOB_DBASE_META_OWN_DELETES_ARE_VISIBLE] = ".ownDeletesAreVisible"
	$asMetaQueries[$LOB_DBASE_META_OWN_INSERTS_ARE_VISIBLE] = ".ownInsertsAreVisible"
	$asMetaQueries[$LOB_DBASE_META_OWN_UPDATES_ARE_VISIBLE] = ".ownUpdatesAreVisible"
	$asMetaQueries[$LOB_DBASE_META_STORES_LOWER_CASE_IDS] = ".storesLowerCaseIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_STORES_MIXED_CASE_IDS] = ".storesMixedCaseIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_STORES_UPPER_CASE_IDS] = ".storesUpperCaseIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_STORES_LOWER_CASE_QUOTED_IDS] = ".storesLowerCaseQuotedIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_STORES_MIXED_CASE_QUOTED_IDS] = ".storesMixedCaseQuotedIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_STORES_UPPER_CASE_QUOTED_IDS] = ".storesUpperCaseQuotedIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ALTER_TABLE_WITH_ADD_COL] = ".supportsAlterTableWithAddColumn"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ALTER_TABLE_WITH_DROP_COL] = ".supportsAlterTableWithDropColumn"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ANSI92_ENTRY_LEVEL_SQL] = ".supportsANSI92EntryLevelSQL"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ANSI92_FULL_SQL] = ".supportsANSI92FullSQL"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ANSI92_INTERMEDIATE_SQL] = ".supportsANSI92IntermediateSQL"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_BATCH_UPDATES] = ".supportsBatchUpdates"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_DATA_MANIPULATION] = ".supportsCatalogsInDataManipulation"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_INDEX_DEFINITIONS] = ".supportsCatalogsInIndexDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_PRIVILEGE_DEFINITIONS] = ".supportsCatalogsInPrivilegeDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_PROCEDURE_CALLS] = ".supportsCatalogsInProcedureCalls"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_TABLE_DEFINITION] = ".supportsCatalogsInTableDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_COL_ALIASING] = ".supportsColumnAliasing"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CONVERT] = ".supportsConvert"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CORE_SQL_GRAMMAR] = ".supportsCoreSQLGrammar"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_CORRELATED_SUBQUERIES] = ".supportsCorrelatedSubqueries"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_DATA_DEFINITION_AND_DATA_MANIPULATION_TRANSACTIONS] = ".supportsDataDefinitionAndDataManipulationTransactions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_DATA_MANIPULATION_TRANSACTIONS_ONLY] = ".supportsDataManipulationTransactionsOnly"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_DIFF_TABLE_CORRELATION_NAMES] = ".supportsDifferentTableCorrelationNames"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_EXPRESSIONS_IN_ORDER_BY] = ".supportsExpressionsInOrderBy"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_EXTENDED_SQL_GRAMMAR] = ".supportsExtendedSQLGrammar"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_FULL_OUTER_JOINS] = ".supportsFullOuterJoins"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_GROUP_BY] = ".supportsGroupBy"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_GROUP_BY_BEYOND_SELECT] = ".supportsGroupByBeyondSelect"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_GROUP_BY_UNRELATED] = ".supportsGroupByUnrelated"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_INTEGRITY_ENHANCMENT_FACILITY] = ".supportsIntegrityEnhancementFacility"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_LIKE_ESCAPE_CLAUSE] = ".supportsLikeEscapeClause"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_LIMITED_OUTER_JOINS] = ".supportsLimitedOuterJoins"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_MINIMUM_SQL_GRAMMAR] = ".supportsMinimumSQLGrammar"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_MIXED_CASE_IDS] = ".supportsMixedCaseIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_MIXED_CASE_QUOTED_IDS] = ".supportsMixedCaseQuotedIdentifiers"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_MULTIPLE_RESULT_SETS] = ".supportsMultipleResultSets"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_MULTIPLE_TRANSACTIONS] = ".supportsMultipleTransactions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_NON_NULLABLE_COLS] = ".supportsNonNullableColumns"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_OPEN_CURSORS_ACROSS_COMMIT] = ".supportsOpenCursorsAcrossCommit"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_OPEN_CURSORS_ACROSS_ROLLBACK] = ".supportsOpenCursorsAcrossRollback"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_OPEN_STATEMENTS_ACROSS_COMMIT] = ".supportsOpenStatementsAcrossCommit"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_OPEN_STATEMENTS_ACROSS_ROLLBACK] = ".supportsOpenStatementsAcrossRollback"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_ORDER_BY_UNRELATED] = ".supportsOrderByUnrelated"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_OUTER_JOINS] = ".supportsOuterJoins"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_POSITIONED_DELETE] = ".supportsPositionedDelete"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_POSITIONED_UPDATE] = ".supportsPositionedUpdate"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_RESULT_SET_CONCURRENCY] = ".supportsResultSetConcurrency"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_RESULT_SET_TYPE] = ".supportsResultSetType"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_DATA_MANIPULATION] = ".supportsSchemasInDataManipulation"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_INDEX_DEFINITIONS] = ".supportsSchemasInIndexDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_PRIVILEGE_DEFINITIONS] = ".supportsSchemasInPrivilegeDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_PROCEDURE_CALLS] = ".supportsSchemasInProcedureCalls"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_TABLE_DEFINITION] = ".supportsSchemasInTableDefinitions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SELECT_FOR_UPDATE] = ".supportsSelectForUpdate"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_STORED_PROCEDURES] = ".supportsStoredProcedures"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_COMPARISONS] = ".supportsSubqueriesInComparisons"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_EXISTS] = ".supportsSubqueriesInExists"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_INS] = ".supportsSubqueriesInIns"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_QUANTIFIEDS] = ".supportsSubqueriesInQuantifieds"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_TABLE_CORRELATION_NAMES] = ".supportsTableCorrelationNames"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_TRANSACTION_ISOLATION_LEVEL] = ".supportsTransactionIsolationLevel"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_TRANSACTIONS] = ".supportsTransactions"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_TYPE_CONVERSION] = ".supportsTypeConversion"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_UNION] = ".supportsUnion"
	$asMetaQueries[$LOB_DBASE_META_SUPPORTS_UNION_ALL] = ".supportsUnionAll"
	$asMetaQueries[$LOB_DBASE_META_UPDATES_ARE_DETECTED] = ".updatesAreDetected"
	$asMetaQueries[$LOB_DBASE_META_USES_LOCAL_FILE_PER_TABLE] = ".usesLocalFilePerTable"
	$asMetaQueries[$LOB_DBASE_META_USES_LOCAL_FILES] = ".usesLocalFiles"

	Return SetError($__LO_STATUS_SUCCESS, 0, $asMetaQueries[$iQuery])
EndFunc   ;==>__LOBase_DatabaseMetaGetQuery

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LOBase_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _LOBase_ComError_UserFunction(Default)
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
EndFunc   ;==>__LOBase_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_ReportConIdentify
; Description ...: Identify the type of Control being called, or return the Service name of a control type.
; Syntax ........: __LOBase_ReportConIdentify($oControl[, $iControlType = Null])
; Parameters ....: $oControl            - an object. A Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
;                  $iControlType        - [optional] an integer value (1-32). Default is Null. The Control Type Constant. See Constants $LOB_REP_CON_TYPE_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: Integer or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object, and $iControlType not an Integer, less than 1 or greater than 32. See Constants $LOB_REP_CON_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control, or return requested Service name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning Constant value for Control type.
;                  @Error 0 @Extended 1 Return String = Success. Returning requested Control type's service name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_ReportConIdentify($oControl, $iControlType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avControls[6][2] = [["com.sun.star.chart2.ChartDocument", $LOB_REP_CON_TYPE_CHART], ["com.sun.star.report.FormattedField", $LOB_REP_CON_TYPE_FORMATTED_FIELD], _
			["com.sun.star.report.ImageControl", $LOB_REP_CON_TYPE_IMAGE_CONTROL], ["com.sun.star.report.FixedText", $LOB_REP_CON_TYPE_LABEL], _
			["com.sun.star.report.FixedLine", $LOB_REP_CON_TYPE_LINE], ["com.sun.star.report.TextField", $LOB_REP_CON_TYPE_TEXT_BOX]]

	If Not IsObj($oControl) And Not __LO_IntIsBetween($iControlType, $LOB_REP_CON_TYPE_CHART, $LOB_REP_CON_TYPE_TEXT_BOX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If IsObj($oControl) Then
		For $i = 0 To UBound($avControls) - 1
			If $oControl.supportsService($avControls[$i][0]) Then Return SetError($__LO_STATUS_SUCCESS, 0, $avControls[$i][1])
		Next

	ElseIf IsInt($iControlType) Then
		For $i = 0 To UBound($avControls) - 1
			If ($avControls[$i][1] = $iControlType) Then Return SetError($__LO_STATUS_SUCCESS, 1, $avControls[$i][0])
		Next
	EndIf

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>__LOBase_ReportConIdentify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOBase_ReportConSetGetFontDesc
; Description ...: Set or Retrieve a Control's Font values.
; Syntax ........: __LOBase_ReportConSetGetFontDesc(ByRef $oControl[, $mFontDesc = Null])
; Parameters ....: $oControl            - [in/out] an object. A Control object returned by a previous _LOBase_ReportConInsert or _LOBase_ReportConsGetList function.
;                  $mFontDesc           - [optional] a map. Default is Null. A Font descriptor Map returned by a previous _LOBase_FontDescCreate or _LOBase_FontDescEdit function.
; Return values .: Success: 1 or Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = $mFontDesc not a Map.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting Font Name.
;                  |                               2 = Error setting Font Weight.
;                  |                               4 = Error setting Font Posture.
;                  |                               8 = Error setting Font Size.
;                  |                               16 = Error setting Font Color.
;                  |                               32 = Error setting Font Underline Style.
;                  |                               64 = Error setting Font Underline Color.
;                  |                               128 = Error setting Font Strikeout Style.
;                  |                               256 = Error setting Individual Word mode.
;                  |                               512 = Error setting Font Relief.
;                  |                               1024 = Error setting Font Case.
;                  |                               2048 = Error setting Font Character Hidden.
;                  |                               4096 = Error setting Font Character Contoured.
;                  |                               8192 = Error setting Font Character Shadowed.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Map = Success. All optional parameters were called with Null, returning current settings as a Map.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOBase_ReportConSetGetFontDesc(ByRef $oControl, $mFontDesc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $mControlFontDesc[]

	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($mFontDesc) Then
		$mControlFontDesc.CharFontName = $oControl.CharFontName()
		$mControlFontDesc.CharWeight = $oControl.CharWeight()
		$mControlFontDesc.CharPosture = $oControl.CharPosture()
		$mControlFontDesc.CharHeight = $oControl.CharHeight()
		$mControlFontDesc.CharColor = $oControl.CharColor()
		$mControlFontDesc.CharUnderline = $oControl.CharUnderline()
		$mControlFontDesc.CharUnderlineColor = $oControl.CharUnderlineColor()
		$mControlFontDesc.CharStrikeout = $oControl.CharStrikeout()
		$mControlFontDesc.CharWordMode = $oControl.CharWordMode()
		$mControlFontDesc.CharRelief = $oControl.CharRelief()
		$mControlFontDesc.CharCaseMap = $oControl.CharCaseMap()
		$mControlFontDesc.CharHidden = $oControl.CharHidden()
		$mControlFontDesc.CharContoured = $oControl.CharContoured()
		$mControlFontDesc.CharShadowed = $oControl.CharShadowed()

		Return SetError($__LO_STATUS_SUCCESS, 1, $mControlFontDesc)
	EndIf

	If Not IsMap($mFontDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oControl.CharFontName() = $mFontDesc.CharFontName
	$iError = ($oControl.CharFontName() = $mFontDesc.CharFontName) ? ($iError) : (BitOR($iError, 1))

	$oControl.CharWeight() = $mFontDesc.CharWeight
	$iError = (__LO_IntIsBetween($oControl.CharWeight(), $mFontDesc.CharWeight - 50, $mFontDesc.CharWeight + 50)) ? ($iError) : (BitOR($iError, 2))

	$oControl.CharPosture() = $mFontDesc.CharPosture
	$iError = ($oControl.CharPosture() = $mFontDesc.CharPosture) ? ($iError) : (BitOR($iError, 4))

	$oControl.CharHeight() = $mFontDesc.CharHeight
	$iError = ($oControl.CharHeight() = $mFontDesc.CharHeight) ? ($iError) : (BitOR($iError, 8))

	$oControl.CharColor() = $mFontDesc.CharColor
	$iError = ($oControl.CharColor() = $mFontDesc.CharColor) ? ($iError) : (BitOR($iError, 16))

	$oControl.CharUnderline() = $mFontDesc.CharUnderline
	$iError = ($oControl.CharUnderline() = $mFontDesc.CharUnderline) ? ($iError) : (BitOR($iError, 32))

	$oControl.CharUnderlineColor() = $mFontDesc.CharUnderlineColor
	$iError = ($oControl.CharUnderlineColor() = $mFontDesc.CharUnderlineColor) ? ($iError) : (BitOR($iError, 64))

	$oControl.CharStrikeout() = $mFontDesc.CharStrikeout
	$iError = ($oControl.CharStrikeout() = $mFontDesc.CharStrikeout) ? ($iError) : (BitOR($iError, 128))

	$oControl.CharWordMode() = $mFontDesc.CharWordMode
	$iError = ($oControl.CharWordMode() = $mFontDesc.CharWordMode) ? ($iError) : (BitOR($iError, 256))

	$oControl.CharRelief() = $mFontDesc.CharRelief
	$iError = ($oControl.CharRelief() = $mFontDesc.CharRelief) ? ($iError) : (BitOR($iError, 512))

	$oControl.CharCaseMap = $mFontDesc.CharCaseMap
	$iError = ($oControl.CharCaseMap() = $mFontDesc.CharCaseMap) ? ($iError) : (BitOR($iError, 1024))

	$oControl.CharHidden = $mFontDesc.CharHidden
	$iError = ($oControl.CharHidden() = $mFontDesc.CharHidden) ? ($iError) : (BitOR($iError, 2048))

	$oControl.CharContoured = $mFontDesc.CharContoured
	$iError = ($oControl.CharContoured() = $mFontDesc.CharContoured) ? ($iError) : (BitOR($iError, 4096))

	$oControl.CharShadowed = $mFontDesc.CharShadowed
	$iError = ($oControl.CharShadowed() = $mFontDesc.CharShadowed) ? ($iError) : (BitOR($iError, 8192))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOBase_ReportConSetGetFontDesc
