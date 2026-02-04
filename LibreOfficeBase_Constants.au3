#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Base Constants for the Libre Office UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various functions in the Libre Office UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; Note ..........: Descriptions for some Constants are taken from the LibreOffice SDK API documentation.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; Sleep Divisor $__LOCCONST_SLEEP_DIV
; In applicable functions this is used for adjusting how frequent a sleep occurs in loops.
; For any number above 0 the number of times a loop has completed is divided by $__LOBCONST_SLEEP_DIV. If you find some functions cause momentary freeze ups, a recommended value is 15.
; Set to 0 for no pause in a loop.
Global Const $__LOBCONST_SLEEP_DIV = 0

#Tidy_ILC_Pos=90

; Vertical Alignment
Global Const _                                                                           ; com.sun.star.style.VerticalAlignment
		$LOB_ALIGN_VERT_TOP = 0, _                                                       ; Vertically Align the object to the Top.
		$LOB_ALIGN_VERT_MIDDLE = 1, _                                                    ; Vertically Align the object to the Middle.
		$LOB_ALIGN_VERT_BOTTOM = 2                                                       ; Vertically Align the object to the Bottom.

; Case Constants
Global Const _                                                                           ; com.sun.star.style.CaseMap
		$LOB_CASEMAP_NONE = 0, _                                                         ; The case of the characters is unchanged.
		$LOB_CASEMAP_UPPER = 1, _                                                        ; All characters are put in upper case.
		$LOB_CASEMAP_LOWER = 2, _                                                        ; All characters are put in lower case.
		$LOB_CASEMAP_TITLE = 3, _                                                        ; The first character of each word is put in upper case.
		$LOB_CASEMAP_SM_CAPS = 4                                                         ; All characters are put in upper case, but with a smaller font height.

; Table Column Text Alignment
Global Const _                                                                           ; com.sun.star.sdb.Align
		$LOB_COL_TXT_ALIGN_LEFT = 0, _                                                   ; The Column's Text is aligned to the Left.
		$LOB_COL_TXT_ALIGN_CENTER = 1, _                                                 ; The Column's Text is aligned in the center.
		$LOB_COL_TXT_ALIGN_RIGHT = 2                                                     ; The Column's Text is aligned to the Right.

; Prepared Statement Input Type Commands.
Global Enum _
		$LOB_DATA_SET_TYPE_NULL, _                                                       ; Sets the content of the column to NULL.
		$LOB_DATA_SET_TYPE_BOOL, _                                                       ; Puts the given logical value into the SQL command.
		$LOB_DATA_SET_TYPE_BYTE, _                                                       ; Puts the given byte into the SQL command.
		$LOB_DATA_SET_TYPE_SHORT, _                                                      ; Puts the given integer into the SQL command.
		$LOB_DATA_SET_TYPE_INT, _                                                        ; Puts the given integer into the SQL command.
		$LOB_DATA_SET_TYPE_LONG, _                                                       ; Puts the given integer into the SQL command.
		$LOB_DATA_SET_TYPE_FLOAT, _                                                      ; Puts the given decimal number into the SQL command.
		$LOB_DATA_SET_TYPE_DOUBLE, _                                                     ; Puts the given decimal number into the SQL command.
		$LOB_DATA_SET_TYPE_STRING, _                                                     ; Puts the given character string into the SQL command.
		$LOB_DATA_SET_TYPE_BYTES, _                                                      ; Puts the given byte array into the SQL command.
		$LOB_DATA_SET_TYPE_DATE, _                                                       ; Puts the given date into the SQL command.
		$LOB_DATA_SET_TYPE_TIME, _                                                       ; Puts the given time into the SQL command.
		$LOB_DATA_SET_TYPE_TIMESTAMP, _                                                  ; Puts the given timestamp into the SQL command.
		$LOB_DATA_SET_TYPE_CLOB, _                                                       ; Puts the given CLOB (Character Large Object) into the SQL command.
		$LOB_DATA_SET_TYPE_BLOB, _                                                       ; Puts the given BLOB (Binary Large Object) into the SQL command.
		$LOB_DATA_SET_TYPE_ARRAY, _                                                      ; Puts the given Array into the SQL command.
		$LOB_DATA_SET_TYPE_OBJECT                                                        ; Puts the given Object into the SQL command.

; Database Data Types
Global Const _                                                                           ; com.sun.star.sdbc.DataType Constant Group
		$LOB_DATA_TYPE_LONGNVARCHAR = -16, _                                             ; L.O. 24.2
		$LOB_DATA_TYPE_NCHAR = -15, _                                                    ; L.O. 24.2
		$LOB_DATA_TYPE_NVARCHAR = -9, _                                                  ; L.O. 24.2
		$LOB_DATA_TYPE_ROWID = -8, _                                                     ; L.O. 24.2
		$LOB_DATA_TYPE_BIT = -7, _
		$LOB_DATA_TYPE_TINYINT = -6, _
		$LOB_DATA_TYPE_BIGINT = -5, _
		$LOB_DATA_TYPE_LONGVARBINARY = -4, _
		$LOB_DATA_TYPE_VARBINARY = -3, _
		$LOB_DATA_TYPE_BINARY = -2, _
		$LOB_DATA_TYPE_LONGVARCHAR = -1, _
		$LOB_DATA_TYPE_SQLNULL = 0, _
		$LOB_DATA_TYPE_CHAR = 1, _
		$LOB_DATA_TYPE_NUMERIC = 2, _
		$LOB_DATA_TYPE_DECIMAL = 3, _
		$LOB_DATA_TYPE_INTEGER = 4, _
		$LOB_DATA_TYPE_SMALLINT = 5, _
		$LOB_DATA_TYPE_FLOAT = 6, _
		$LOB_DATA_TYPE_REAL = 7, _
		$LOB_DATA_TYPE_DOUBLE = 8, _
		$LOB_DATA_TYPE_VARCHAR = 12, _
		$LOB_DATA_TYPE_BOOLEAN = 16, _
		$LOB_DATA_TYPE_DATALINK = 70, _                                                  ; L.O. 24.2
		$LOB_DATA_TYPE_DATE = 91, _
		$LOB_DATA_TYPE_TIME = 92, _
		$LOB_DATA_TYPE_TIMESTAMP = 93, _
		$LOB_DATA_TYPE_OTHER = 1111, _
		$LOB_DATA_TYPE_OBJECT = 2000, _
		$LOB_DATA_TYPE_DISTINCT = 2001, _
		$LOB_DATA_TYPE_STRUCT = 2002, _
		$LOB_DATA_TYPE_ARRAY = 2003, _
		$LOB_DATA_TYPE_BLOB = 2004, _
		$LOB_DATA_TYPE_CLOB = 2005, _
		$LOB_DATA_TYPE_REF = 2006, _
		$LOB_DATA_TYPE_SQLXML = 2009, _                                                  ; L.O. 24.2
		$LOB_DATA_TYPE_NCLOB = 2011, _                                                   ; L.O. 24.2
		$LOB_DATA_TYPE_REF_CURSOR = 2012, _                                              ; L.O. 24.2
		$LOB_DATA_TYPE_TIME_WITH_TIMEZONE = 2013, _                                      ; L.O. 24.2
		$LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE = 2014                                    ; L.O. 24.2

; Best Row Scope.
Global Const _                                                                           ; "com.sun.star.sdbc.BestRowScope"
		$LOB_DBASE_BEST_ROW_SCOPE_TEMPORARY = 0, _                                       ; Indicates that the scope of the best row identifier is very temporary, lasting only while the row is being used.
		$LOB_DBASE_BEST_ROW_SCOPE_TRANSACTION = 1, _                                     ; Indicates that the scope of the best row identifier is the remainder of the current transaction.
		$LOB_DBASE_BEST_ROW_SCOPE_SESSION = 2                                            ; Indicates that the scope of the best row identifier is the remainder of the current session.

; Database MetaData Queries
Global Enum _                                                                            ; Based on "com.sun.star.sdbc.XDatabaseMetaData"
		$LOB_DBASE_META_ALL_PROCEDURES_ARE_CALLABLE, _                                   ; Returns a Boolean. Can all the procedures returned by getProcedures be called by the current user? TRUE if the user is allowed to call all procedures returned by getProcedures otherwise FALSE.
		$LOB_DBASE_META_ALL_TABLES_ARE_SELECTABLE, _                                     ; Returns a Boolean. Can all the tables returned by getTable be SELECTed by the current user? True if so.
		$LOB_DBASE_META_DATA_DEFINITION_CAUSES_TRANSACTION_COMMIT, _                     ; Returns a Boolean. Does a data definition statement within a transaction force the transaction to commit? True if so.
		$LOB_DBASE_META_DATA_DEFINITION_IGNORED_IN_TRANSACTIONS, _                       ; Returns a Boolean. Is a data definition statement within a transaction ignored? True if so.
		$LOB_DBASE_META_DELETES_ARE_DETECTED, _                                          ; Returns a Boolean. Indicates whether or not a visible row delete can be detected by calling $LOB_RESULT_ROW_QUERY_IS_ROW_DELETED. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_DOES_MAX_ROW_SIZE_INCLUDE_BLOBS, _                               ; Did $LOB_DBASE_META_GET_MAX_ROW_SIZE include LONGVARCHAR and LONGVARBINARY blobs? Returns a Boolean. True if so.
		$LOB_DBASE_META_GET_BEST_ROW_ID, _                                               ; Returns a Result Set, each row is a column description. Gets a description of a table's optimal set of columns that uniquely identifies a row. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_CATALOG_SEPARATOR, _                                         ; Returns a String. Return the separator between catalog and table name.
		$LOB_DBASE_META_GET_CATALOG_TERM, _                                              ; Returns a String. Return the database vendor's preferred term for "catalog".
		$LOB_DBASE_META_GET_CATALOGS, _                                                  ; Returns a Result Set, each row has a single String column that is a catalog name. Gets the catalog names available in this database. The catalog column is: TABLE_CAT string => catalog name.
		$LOB_DBASE_META_GET_COLS, _                                                      ; Returns a Result Set, each row is a column privilege description. Gets a description of table columns available in the specified catalog. Only column descriptions matching the catalog, schema, table and column name criteria are returned. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_COL_PRIVILEGES, _                                            ; Returns a Result Set, each row is a column description. Gets a description of the access rights for a table's columns. Only privileges matching the column name criteria are returned. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_CROSS_REF, _                                                 ; Returns a Result Set, each row is a foreign key column description. Gets a description of the foreign key columns in the foreign key table that reference the primary key columns of the primary key table (describe how one table imports another's key.) This should normally return a single foreign key/primary key pair (most tables only import a foreign key from a table once.). See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_DATABASE_PRODUCT_NAME, _                                     ; Returns a String. Returns the name of the database product.
		$LOB_DBASE_META_GET_DATABASE_PRODUCT_VERSION, _                                  ; Returns a String. Returns the version of the database product.
		$LOB_DBASE_META_GET_DEFAULT_TRANSACTION_ISOLATION, _                             ; Returns an Integer. Return the database default transaction isolation level. See $LOB_DBASE_TRANSACTION_ISOLATION_* Constants.
		$LOB_DBASE_META_GET_DRIVER_MAJOR_VERSION, _                                      ; Returns an Integer. Returns the SDBC driver major version number.
		$LOB_DBASE_META_GET_DRIVER_MINOR_VERSION, _                                      ; Returns an Integer. Returns the SDBC driver minor version number.
		$LOB_DBASE_META_GET_DRIVER_NAME, _                                               ; Returns a String. Returns the name of the SDBC driver.
		$LOB_DBASE_META_GET_DRIVER_VERSION, _                                            ; Returns an Integer. Returns the version number of the SDBC driver.
		$LOB_DBASE_META_GET_EXPORTED_KEYS, _                                             ; Returns a Result Set, each row is a foreign key column description. Gets a description of the foreign key columns that reference a table's primary key columns (the foreign keys exported by a table). See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_EXTRA_NAME_CHARS, _                                          ; Returns a String. Gets all the "extra" characters that can be used in unquoted identifier names (those beyond a-z, A-Z, 0-9 and _).
		$LOB_DBASE_META_GET_IDENTIFIER_QUOTE_STRING, _                                   ; Returns a String. What's the string used to quote SQL identifiers? This returns a space " " if identifier quoting is not supported.
		$LOB_DBASE_META_GET_IMPORTED_KEYS, _                                             ; Returns a Result Set, each row is a primary key column description. Gets a description of the primary key columns that are referenced by a table's foreign key columns (the primary keys imported by a table). See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_INDEX_INFO, _                                                ; Returns a Result Set, each row is an index column description. Gets a description of a table's indices and statistics. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_MAX_BINARY_LITERAL_LEN, _                                    ; Returns an Integer. Return the maximal number of hex characters in an inline binary literal.
		$LOB_DBASE_META_GET_MAX_CATALOG_NAME_LEN, _                                      ; Returns an Integer. Return the maximum length of a catalog name.
		$LOB_DBASE_META_GET_MAX_CHAR_LITERAL_LEN, _                                      ; Returns an Integer. Return the max length for a character literal.
		$LOB_DBASE_META_GET_MAX_COL_NAME_LEN, _                                          ; Returns an Integer. Return the limit on column name length.
		$LOB_DBASE_META_GET_MAX_COLS_IN_GROUP_BY, _                                      ; Returns an Integer. Return the maximum number of columns in a "GROUP BY" clause.
		$LOB_DBASE_META_GET_MAX_COLS_IN_INDEX, _                                         ; Returns an Integer. Return the maximum number of columns allowed in an index.
		$LOB_DBASE_META_GET_MAX_COLS_IN_ORDER_BY, _                                      ; Returns an Integer. Return the maximum number of columns in an "ORDER BY" clause.
		$LOB_DBASE_META_GET_MAX_COLS_IN_SEL, _                                           ; Returns an Integer. Return the maximum number of columns in a "SELECT" list.
		$LOB_DBASE_META_GET_MAX_COLS_IN_TABLE, _                                         ; Returns an Integer. Return the maximum number of columns in a table.
		$LOB_DBASE_META_GET_MAX_CONNECTIONS, _                                           ; Returns an Integer. Return the number of active connections at a time to this database.
		$LOB_DBASE_META_GET_MAX_CURSOR_NAME_LEN, _                                       ; Returns an Integer. Return the maximum cursor name length.
		$LOB_DBASE_META_GET_MAX_INDEX_LEN, _                                             ; Returns an Integer. Return the maximum length of an index (in bytes).
		$LOB_DBASE_META_GET_MAX_PROCEDURE_NAME_LEN, _                                    ; Returns an Integer. Return the maximum length of a procedure name.
		$LOB_DBASE_META_GET_MAX_ROW_SIZE, _                                              ; Returns an Integer. Return the maximum length of a single row.
		$LOB_DBASE_META_GET_MAX_SCHEMA_NAME_LEN, _                                       ; Returns an Integer. Return the maximum length allowed for a schema name.
		$LOB_DBASE_META_GET_MAX_STATEMENT_LEN, _                                         ; Returns an Integer. Return the maximum length of a SQL statement.
		$LOB_DBASE_META_GET_MAX_STATEMENTS, _                                            ; Returns an Integer. Return the maximal number of open active statements at one time to this database.
		$LOB_DBASE_META_GET_MAX_TABLE_NAME_LEN, _                                        ; Returns an Integer. Return the maximum length of a table name.
		$LOB_DBASE_META_GET_MAX_TABLES_IN_SEL, _                                         ; Returns an Integer. Return the maximum number of tables in a SELECT statement.
		$LOB_DBASE_META_GET_MAX_USER_NAME_LEN, _                                         ; Returns an Integer. Return the maximum length of a user name.
		$LOB_DBASE_META_GET_NUMERIC_FUNCS, _                                             ; Returns a String. Gets a comma-separated list of math functions. These are the X/Open CLI math function names used in the SDBC function escape clause.
		$LOB_DBASE_META_GET_PRIMARY_KEY, _                                               ; Returns a Result Set, each row is a primary key column description. Gets a description of a table's primary key columns. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_PROCEDURE_COLS, _                                            ; Returns a Result Set, each row describes a stored procedure parameter or column. Gets a description of a catalog's stored procedure parameters and result columns. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_PROCEDURE_TERM, _                                            ; Returns a String. Return the database vendor's preferred term for "procedure".
		$LOB_DBASE_META_GET_PROCEDURES, _                                                ; Returns a Result Set, each row is a procedure description. Gets a description of the stored procedures available in a catalog. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_SCHEMA_TERM, _                                               ; Returns a String. Return the database vendor's preferred term for "schema".
		$LOB_DBASE_META_GET_SCHEMAS, _                                                   ; Returns a Result Set, each row has a single String column that is a schema name. Gets the schema names available in this database. The schema column is: TABLE_SCHEM string => schema name.
		$LOB_DBASE_META_GET_SEARCH_STRING_ESCAPE, _                                      ; Returns a String. Gets the string that can be used to escape wildcard characters. This is the string that can be used to escape "_" or "%" in the string pattern style catalog search parameters.
		$LOB_DBASE_META_GET_SQL_KEYWORDS, _                                              ; Returns a String. Gets a comma-separated list of all a database's SQL keywords that are NOT also SQL92 keywords.
		$LOB_DBASE_META_GET_STRING_FUNCS, _                                              ; Returns a String. Gets a comma-separated list of string functions. These are the X/Open CLI string function names used in the SDBC function escape clause.
		$LOB_DBASE_META_GET_SYSTEM_FUNCS, _                                              ; Returns a String. Gets a comma-separated list of system functions. These are the X/Open CLI system function names used in the SDBC function escape clause.
		$LOB_DBASE_META_GET_TABLE_PRIVILEGES, _                                          ; Returns a Result Set, each row is a table privilege description. Gets a description of the access rights for each table available in a catalog. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_TABLE_TYPES, _                                               ; Returns a Result Set, each row has a single String column that is a table type. Gets the table types available in this database.  See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_TABLES, _                                                    ; Returns a Result Set, each row is a table description.  Gets a description of tables available in a catalog. See _LOBase_DatabaseMetaDataQuery for Parameters.See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_TIME_DATE_FUNCS, _                                           ; Returns a String. Gets a comma-separated list of time and date functions.
		$LOB_DBASE_META_GET_TYPE_INFO, _                                                 ; Returns a Result Set, each row is a SQL type description. Gets a description of all the standard SQL types supported by this database. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_UDTS, _                                                      ; Returns a Result Set, each row is a type description. Gets a description of the user-defined types defined in a particular schema. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_VERSION_COLS, _                                              ; Returns a Result Set, each row is a column description. Gets a description of a table's columns that are automatically updated when any value in a row is updated. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_URL, _                                                       ; Returns a String. Returns the URL for the database connection.
		$LOB_DBASE_META_GET_USERNAME, _                                                  ; Returns a String. Returns the user name from this database connection.
		$LOB_DBASE_META_INSERTS_ARE_DETECTED, _                                          ; Returns a Boolean. Indicates whether or not a visible row insert can be detected by calling $LOB_RESULT_ROW_QUERY_IS_ROW_INSERTED. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_IS_CATALOG_AT_START, _                                           ; Returns a Boolean. Does a catalog appear at the start of a qualified table name? (Otherwise it appears at the end). True if so.
		$LOB_DBASE_META_IS_READ_ONLY, _                                                  ; Returns a Boolean. Checks if the database in read-only mode. True if so.
		$LOB_DBASE_META_NULL_PLUS_NON_NULL_IS_NULL, _                                    ; Returns a Boolean. Are concatenations between NULL and non-NULL values NULL? True if so.
		$LOB_DBASE_META_NULLS_ARE_SORTED_AT_END, _                                       ; Returns a Boolean. Are NULL values sorted at the end, regardless of sort order? True if so.
		$LOB_DBASE_META_NULLS_ARE_SORTED_AT_START, _                                     ; Returns a Boolean. Are NULL values sorted at the start regardless of sort order? True if so.
		$LOB_DBASE_META_NULLS_ARE_SORTED_HIGH, _                                         ; Returns a Boolean. Are NULL values sorted high? True if so.
		$LOB_DBASE_META_NULLS_ARE_SORTED_LOW, _                                          ; Returns a Boolean. Are NULL values sorted low? True if so.
		$LOB_DBASE_META_OTHERS_DELETES_ARE_VISIBLE, _                                    ; Returns a Boolean. Indicates whether deletes made by others are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_OTHERS_INSERTS_ARE_VISIBLE, _                                    ; Returns a Boolean. Indicates whether inserts made by others are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_OTHERS_UPDATES_ARE_VISIBLE, _                                    ; Returns a Boolean. Indicates whether updates made by others are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_OWN_DELETES_ARE_VISIBLE, _                                       ; Returns a Boolean. Indicates whether a result set's own deletes are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_OWN_INSERTS_ARE_VISIBLE, _                                       ; Returns a Boolean. Indicates whether a result set's own inserts are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_OWN_UPDATES_ARE_VISIBLE, _                                       ; Returns a Boolean. Indicates whether a result set's own updates are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_STORES_LOWER_CASE_IDS, _                                         ; Returns a Boolean. Does the database treat mixed case unquoted SQL identifiers as case insensitive and store them in lower case? True if so.
		$LOB_DBASE_META_STORES_MIXED_CASE_IDS, _                                         ; Returns a Boolean. Does the database treat mixed case unquoted SQL identifiers as case insensitive and store them in mixed case? True if so.
		$LOB_DBASE_META_STORES_UPPER_CASE_IDS, _                                         ; Returns a Boolean. Does the database treat mixed case unquoted SQL identifiers as case insensitive and store them in upper case? True if so.
		$LOB_DBASE_META_STORES_LOWER_CASE_QUOTED_IDS, _                                  ; Returns a Boolean. Does the database treat mixed case quoted SQL identifiers as case insensitive and store them in lower case? True if so.
		$LOB_DBASE_META_STORES_MIXED_CASE_QUOTED_IDS, _                                  ; Returns a Boolean. Does the database treat mixed case quoted SQL identifiers as case insensitive and store them in mixed case? True if so.
		$LOB_DBASE_META_STORES_UPPER_CASE_QUOTED_IDS, _                                  ; Returns a Boolean. Does the database treat mixed case quoted SQL identifiers as case insensitive and store them in upper case? True if so.
		$LOB_DBASE_META_SUPPORTS_ALTER_TABLE_WITH_ADD_COL, _                             ; Returns a Boolean. Support the Database "ALTER TABLE" with add column? True if so.
		$LOB_DBASE_META_SUPPORTS_ALTER_TABLE_WITH_DROP_COL, _                            ; Returns a Boolean. Support the Database "ALTER TABLE" with drop column? True if so.
		$LOB_DBASE_META_SUPPORTS_ANSI92_ENTRY_LEVEL_SQL, _                               ; Returns a Boolean. TRUE, if the database supports ANSI92 entry level SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_ANSI92_FULL_SQL, _                                      ; Returns a Boolean. TRUE, if the database supports ANSI92 full SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_ANSI92_INTERMEDIATE_SQL, _                              ; Returns a Boolean. TRUE, if the database supports ANSI92 intermediate SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_BATCH_UPDATES, _                                        ; Returns a Boolean. Indicates whether the driver supports batch updates. True if so.
		$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_DATA_MANIPULATION, _                        ; Returns a Boolean. Can a catalog name be used in a data manipulation statement? True if so.
		$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_INDEX_DEFINITIONS, _                        ; Returns a Boolean. Can a catalog name be used in an index definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_PRIVILEGE_DEFINITIONS, _                    ; Returns a Boolean. Can a catalog name be used in a privilege definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_PROCEDURE_CALLS, _                          ; Returns a Boolean. Can a catalog name be used in a procedure call statement? True if so.
		$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_TABLE_DEFINITION, _                         ; Returns a Boolean. Can a catalog name be used in a table definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_COL_ALIASING, _                                         ; Returns a Boolean. Support the Database column aliasing? The SQL AS clause can be used to provide names for computed columns or to provide alias names for columns as required. True if so.
		$LOB_DBASE_META_SUPPORTS_CONVERT, _                                              ; Returns a Boolean. TRUE , if the Database supports the CONVERT between the given SQL types otherwise FALSE. See _LOBase_DatabaseMetaDataQuery for Parameters.
		$LOB_DBASE_META_SUPPORTS_CORE_SQL_GRAMMAR, _                                     ; Returns a Boolean. TRUE, if the database supports ODBC Core SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_CORRELATED_SUBQUERIES, _                                ; Returns a Boolean. Are correlated subqueries supported? True if so.
		$LOB_DBASE_META_SUPPORTS_DATA_DEFINITION_AND_DATA_MANIPULATION_TRANSACTIONS, _   ; Returns a Boolean. Support the Database both data definition and data manipulation statements within a transaction? True if so.
		$LOB_DBASE_META_SUPPORTS_DATA_MANIPULATION_TRANSACTIONS_ONLY, _                  ; Returns a Boolean. Are only data manipulation statements within a transaction supported? True if so.
		$LOB_DBASE_META_SUPPORTS_DIFF_TABLE_CORRELATION_NAMES, _                         ; Returns a Boolean. If table correlation names are supported, are they restricted to be different from the names of the tables? True if so.
		$LOB_DBASE_META_SUPPORTS_EXPRESSIONS_IN_ORDER_BY, _                              ; Returns a Boolean. Are expressions in "ORDER BY" lists supported? True if so.
		$LOB_DBASE_META_SUPPORTS_EXTENDED_SQL_GRAMMAR, _                                 ; Returns a Boolean. TRUE, if the database supports ODBC Extended SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_FULL_OUTER_JOINS, _                                     ; Returns a Boolean. TRUE, if full nested outer joins are supported, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_GROUP_BY, _                                             ; Returns a Boolean. Is some form of "GROUP BY" clause supported? True if so.
		$LOB_DBASE_META_SUPPORTS_GROUP_BY_BEYOND_SELECT, _                               ; Returns a Boolean. Can a "GROUP BY" clause add columns not in the SELECT provided it specifies all the columns in the SELECT? True if so.
		$LOB_DBASE_META_SUPPORTS_GROUP_BY_UNRELATED, _                                   ; Returns a Boolean. Can a "GROUP BY" clause use columns not in the SELECT? True if so.
		$LOB_DBASE_META_SUPPORTS_INTEGRITY_ENHANCMENT_FACILITY, _                        ; Returns a Boolean. TRUE, if the Database supports SQL Integrity Enhancement Facility, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_LIKE_ESCAPE_CLAUSE, _                                   ; Returns a Boolean. Is the escape character in "LIKE" clauses supported? True if so.
		$LOB_DBASE_META_SUPPORTS_LIMITED_OUTER_JOINS, _                                  ; Returns a Boolean. TRUE, if there is limited support for outer joins. (This will be TRUE if supportFullOuterJoins is TRUE.) FALSE is returned otherwise.
		$LOB_DBASE_META_SUPPORTS_MINIMUM_SQL_GRAMMAR, _                                  ; Returns a Boolean. TRUE, if the database supports ODBC Minimum SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_MIXED_CASE_IDS, _                                       ; Returns a Boolean. Use the database "mixed case unquoted SQL identifiers" case sensitive. True if so.
		$LOB_DBASE_META_SUPPORTS_MIXED_CASE_QUOTED_IDS, _                                ; Returns a Boolean.  Does the database treat mixed case quoted SQL identifiers as case sensitive and as a result store them in mixed case?True if so.
		$LOB_DBASE_META_SUPPORTS_MULTIPLE_RESULT_SETS, _                                 ; Returns a Boolean. Are multiple XResultSets from a single execute supported? True if so.
		$LOB_DBASE_META_SUPPORTS_MULTIPLE_TRANSACTIONS, _                                ; Returns a Boolean. Can we have multiple transactions open at once (on different connections)? True if so.
		$LOB_DBASE_META_SUPPORTS_NON_NULLABLE_COLS, _                                    ; Returns a Boolean. Can columns be defined as non-nullable? True if so.
		$LOB_DBASE_META_SUPPORTS_OPEN_CURSORS_ACROSS_COMMIT, _                           ; Returns a Boolean. Can cursors remain open across commits? True if so.
		$LOB_DBASE_META_SUPPORTS_OPEN_CURSORS_ACROSS_ROLLBACK, _                         ; Returns a Boolean. Can cursors remain open across rollbacks? True if so.
		$LOB_DBASE_META_SUPPORTS_OPEN_STATEMENTS_ACROSS_COMMIT, _                        ; Returns a Boolean. Can statements remain open across commits? True if so.
		$LOB_DBASE_META_SUPPORTS_OPEN_STATEMENTS_ACROSS_ROLLBACK, _                      ; Returns a Boolean. Can statements remain open across rollbacks? True if so.
		$LOB_DBASE_META_SUPPORTS_ORDER_BY_UNRELATED, _                                   ; Returns a Boolean. Can an "ORDER BY" clause use columns not in the SELECT statement? True if so.
		$LOB_DBASE_META_SUPPORTS_OUTER_JOINS, _                                          ; Returns a Boolean. TRUE, if some form of outer join is supported, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_POSITIONED_DELETE, _                                    ; Returns a Boolean. Is positioned DELETE supported? True if so.
		$LOB_DBASE_META_SUPPORTS_POSITIONED_UPDATE, _                                    ; Returns a Boolean. Is positioned UPDATE supported? True if so.
		$LOB_DBASE_META_SUPPORTS_RESULT_SET_CONCURRENCY, _                               ; Returns a Boolean. Does the database support the concurrency type in combination with the given result set type? See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_SUPPORTS_RESULT_SET_TYPE, _                                      ; Returns a Boolean. Does the database support the given result set type? See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_DATA_MANIPULATION, _                         ; Returns a Boolean. Can a schema name be used in a data manipulation statement? True if so.
		$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_INDEX_DEFINITIONS, _                         ; Returns a Boolean. Can a schema name be used in an index definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_PRIVILEGE_DEFINITIONS, _                     ; Returns a Boolean. Can a schema name be used in a privilege definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_PROCEDURE_CALLS, _                           ; Returns a Boolean. Can a schema name be used in a procedure call statement? True if so.
		$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_TABLE_DEFINITION, _                          ; Returns a Boolean. Can a schema name be used in a table definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_SELECT_FOR_UPDATE, _                                    ; Returns a Boolean. Is SELECT for UPDATE supported? True if so.
		$LOB_DBASE_META_SUPPORTS_STORED_PROCEDURES, _                                    ; Returns a Boolean. Are stored procedure calls using the stored procedure escape syntax supported? True if so.
		$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_COMPARISONS, _                            ; Returns a Boolean. Are subqueries in comparison expressions supported? True if so.
		$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_EXISTS, _                                 ; Returns a Boolean. Are subqueries in "exists" expressions supported? True if so.
		$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_INS, _                                    ; Returns a Boolean. Are subqueries in "in" statements supported? True if so.
		$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_QUANTIFIEDS, _                            ; Returns a Boolean. Are subqueries in quantified expressions supported? True if so.
		$LOB_DBASE_META_SUPPORTS_TABLE_CORRELATION_NAMES, _                              ; Returns a Boolean. Are table correlation names supported? True if so.
		$LOB_DBASE_META_SUPPORTS_TRANSACTION_ISOLATION_LEVEL, _                          ; Returns a Boolean. Does this database support the given transaction isolation level? See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_SUPPORTS_TRANSACTIONS, _                                         ; Returns a Boolean. Support the Database transactions? If not, invoking the method com::sun::star::sdbc::XConnection::commit() is a noop and the isolation level is TransactionIsolation_NONE. True if so.
		$LOB_DBASE_META_SUPPORTS_TYPE_CONVERSION, _                                      ; Returns a Boolean. TRUE , if the Database supports the CONVERT function between SQL types, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_UNION, _                                                ; Returns a Boolean. Is SQL UNION supported? True if so.
		$LOB_DBASE_META_SUPPORTS_UNION_ALL, _                                            ; Returns a Boolean. Is SQL UNION ALL supported? True if so.
		$LOB_DBASE_META_UPDATES_ARE_DETECTED, _                                          ; Returns a Boolean. Indicates whether or not a visible row update can be detected by calling $LOB_RESULT_ROW_QUERY_IS_ROW_UPDATED. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_USES_LOCAL_FILE_PER_TABLE, _                                     ; Returns a Boolean. Use the database one local file to save for each table. True if so.
		$LOB_DBASE_META_USES_LOCAL_FILES                                                 ; Returns a Boolean. Use the database local files to save the tables. True if so.

; Result Set Concurrency
Global Const _                                                                           ; "com.sun.star.sdbc.ResultSetConcurrency"
		$LOB_DBASE_RESULT_SET_CONCURRENCY_READ_ONLY = 1007, _                            ; Is the concurrency mode for a Result Set object that may NOT be updated.
		$LOB_DBASE_RESULT_SET_CONCURRENCY_UPDATABLE = 1008                               ; Is the concurrency mode for a Result Set object that may be updated.

Global Const _                                                                           ; com.sun.star.sdbc.TransactionIsolation
		$LOB_DBASE_TRANSACTION_ISOLATION_NONE = 0, _                                     ; Indicates that transactions are not supported.
		$LOB_DBASE_TRANSACTION_ISOLATION_READ_UNCOMMITTED = 1, _                         ; Dirty reads, non-repeatable reads and phantom reads can occur. This level allows a row changed by one transaction to be read by another transaction before any changes in that row have been committed (a "dirty read"). If any of the changes are rolled back, the second transaction will have retrieved an invalid row.
		$LOB_DBASE_TRANSACTION_ISOLATION_READ_COMMITTED = 2, _                           ; Dirty reads are prevented; non-repeatable reads and phantom reads can occur. This level only prohibits a transaction from reading a row with uncommitted changes in it.
		$LOB_DBASE_TRANSACTION_ISOLATION_REPEATABLE_READ = 4, _                          ; Dirty reads and non-repeatable reads are prevented; phantom reads can occur. This level prohibits a transaction from reading a row with uncommitted changes in it, and it also prohibits the situation where one transaction reads a row, a second transaction alters the row, and the first transaction rereads the row, getting different values the second time (a "non-repeatable read").
		$LOB_DBASE_TRANSACTION_ISOLATION_SERIALIZED = 8                                  ; Dirty reads, non-repeatable reads and phantom reads are prevented. This level includes the prohibitions in REPEATABLE_READ and further prohibits the situation where one transaction reads all rows that satisfy a WHERE condition, a second transaction inserts a row that satisfies that WHERE condition, and the first transaction rereads for the same condition, retrieving the additional "phantom" row in the second read.

; Format Key Type
Global Const _                                                                           ; com.sun.star.util.NumberFormat
		$LOB_FORMAT_KEYS_ALL = 0, _                                                      ; Returns All number formats.
		$LOB_FORMAT_KEYS_DEFINED = 1, _                                                  ; Returns Only user-defined number formats.
		$LOB_FORMAT_KEYS_DATE = 2, _                                                     ; Returns Date formats.
		$LOB_FORMAT_KEYS_TIME = 4, _                                                     ; Returns Time formats.
		$LOB_FORMAT_KEYS_DATE_TIME = 6, _                                                ; Returns Number formats which contain date and time.
		$LOB_FORMAT_KEYS_CURRENCY = 8, _                                                 ; Returns Currency formats.
		$LOB_FORMAT_KEYS_NUMBER = 16, _                                                  ; Returns Decimal number formats.
		$LOB_FORMAT_KEYS_SCIENTIFIC = 32, _                                              ; Returns Scientific number formats.
		$LOB_FORMAT_KEYS_FRACTION = 64, _                                                ; Returns Number formats for fractions.
		$LOB_FORMAT_KEYS_PERCENT = 128, _                                                ; Returns Percentage number formats.
		$LOB_FORMAT_KEYS_TEXT = 256, _                                                   ; Returns Text number formats.
		$LOB_FORMAT_KEYS_LOGICAL = 1024, _                                               ; Returns Boolean number formats.
		$LOB_FORMAT_KEYS_UNDEFINED = 2048, _                                             ; Returns Is used as a return value if no format exists.
		$LOB_FORMAT_KEYS_EMPTY = 4096, _                                                 ; Returns Empty Number formats (?)
		$LOB_FORMAT_KEYS_DURATION = 8196                                                 ; Returns Duration number formats.

; Posture/Italic
Global Const _                                                                           ; com.sun.star.awt.FontSlant
		$LOB_POSTURE_NONE = 0, _                                                         ; Specifies a font without slant.
		$LOB_POSTURE_OBLIQUE = 1, _                                                      ; Specifies an oblique font (slant not designed into the font).
		$LOB_POSTURE_ITALIC = 2, _                                                       ; Specifies an italic font (slant designed into the font).
		$LOB_POSTURE_DONTKNOW = 3, _                                                     ; Specifies a font with an unknown slant. For Read Only.
		$LOB_POSTURE_REV_OBLIQUE = 4, _                                                  ; Specifies a reverse oblique font (slant not designed into the font).
		$LOB_POSTURE_REV_ITALIC = 5                                                      ; Specifies a reverse italic font (slant designed into the font).

; Relief
Global Const _                                                                           ; com.sun.star.text.FontRelief
		$LOB_RELIEF_NONE = 0, _                                                          ; No relief is applied.
		$LOB_RELIEF_EMBOSSED = 1, _                                                      ; The font relief is embossed.
		$LOB_RELIEF_ENGRAVED = 2                                                         ; The font relief is engraved.

; Report Control Image Scale.
Global Const _                                                                           ; "com.sun.star.awt.ImageScaleMode"
		$LOB_REP_CON_IMG_BTN_SCALE_NONE = 0, _                                           ; No scaling should happen at all.
		$LOB_REP_CON_IMG_BTN_SCALE_KEEP_ASPECT = 1, _                                    ; The image should be scaled up or down to the size of the surrounding area by keeping its aspect ratio.
		$LOB_REP_CON_IMG_BTN_SCALE_FIT = 2                                               ; The image should be scaled up or down to the size of the surrounding area, he image will finally cover all of the surrounding area, but its dimensions might be distorted.

; Report Line Orientation.
Global Const _                                                                           ; com.sun.star.report.XFixedLine.Orientation
		$LOB_REP_CON_LINE_HORI = 0, _                                                    ; The line is Horizontally aligned.
		$LOB_REP_CON_LINE_VERT = 1                                                       ; The line is Vertically aligned.

; Report Control Type.
Global Enum Step *2 _
		$LOB_REP_CON_TYPE_CHART = 1, _                                                   ; A Chart
		$LOB_REP_CON_TYPE_FORMATTED_FIELD, _                                             ; FORMATTED_FIELD
		$LOB_REP_CON_TYPE_IMAGE_CONTROL, _                                               ; IMAGE_CONTROL
		$LOB_REP_CON_TYPE_LABEL, _                                                       ; FIXED_TEXT
		$LOB_REP_CON_TYPE_LINE, _                                                        ; A Vertical or Horizontal line.
		$LOB_REP_CON_TYPE_TEXT_BOX, _                                                    ; A Text Box.
		$LOB_REP_CON_TYPE_ALL = 63                                                       ; All of the above Control Types. (This value is the BitOR value of all above)

; Report Content Type.
Global Const _                                                                           ; com.sun.star.sdb.CommandType
		$LOB_REP_CONTENT_TYPE_TABLE = 0, _                                               ; The Content Type is a Table.
		$LOB_REP_CONTENT_TYPE_QUERY = 1, _                                               ; The Content Type is a Query.
		$LOB_REP_CONTENT_TYPE_SQL = 2                                                    ; The Content Type is a SQL Command.

; Force New Page Constants.
Global Const _                                                                           ; com.sun.star.report.ForceNewPage
		$LOB_REP_FORCE_PAGE_NONE = 0, _                                                  ; The current section is printed on the current page.
		$LOB_REP_FORCE_PAGE_BEFORE_SECTION = 1, _                                        ; The current section is printed at the top of a new page.
		$LOB_REP_FORCE_PAGE_AFTER_SECTION = 2, _                                         ; The next section following the current section is printed at the top of a new page.
		$LOB_REP_FORCE_PAGE_BEFORE_AFTER_SECTION = 3                                     ; The current section is printed at the top of a new page as well as the next section.

; Report Data Grouping Constants
Global Const _                                                                           ; com.sun.star.report.GroupOn
		$LOB_REP_GROUP_ON_DEFAULT = 0, _                                                 ; The same value in the column value or expression.
		$LOB_REP_GROUP_ON_PREFIX = 1, _                                                  ; The same first nth of characters in the column value or expression.
		$LOB_REP_GROUP_ON_YEAR = 2, _                                                    ; Dates in the same calendar year.
		$LOB_REP_GROUP_ON_QUARTAL = 3, _                                                 ; Dates in the same calendar quarter.
		$LOB_REP_GROUP_ON_MONTH = 4, _                                                   ; Dates in the same month.
		$LOB_REP_GROUP_ON_WEEK = 5, _                                                    ; Dates in the same week.
		$LOB_REP_GROUP_ON_DAY = 6, _                                                     ; Dates on the same day.
		$LOB_REP_GROUP_ON_HOUR = 7, _                                                    ; Times in the same hour.
		$LOB_REP_GROUP_ON_MINUTE = 8, _                                                  ; Times in the same minute.
		$LOB_REP_GROUP_ON_INTERVAL = 9                                                   ; Values within an interval you specify.

; Report Keep Together Constants.
Global Const _                                                                           ; com.sun.star.report.KeepTogether
		$LOB_REP_KEEP_TOG_NO = 0, _                                                      ; Prints the group without keeping the header, detail, and footer together on the same page.
		$LOB_REP_KEEP_TOG_WHOLE_GROUP = 1, _                                             ; Prints the group header, detail, and footer together on the same page.
		$LOB_REP_KEEP_TOG_WITH_FIRST_DETAIL = 2                                          ; Prints the group header on a page when the first detail record can fit on the same page.

; Report Output Document Type.
Global Enum _
		$LOB_REP_OUTPUT_TYPE_UNKNOWN, _                                                  ; The Output Document when the Report is executed is a unknown type.
		$LOB_REP_OUTPUT_TYPE_TEXT, _                                                     ; The Output Document when the Report is executed will be a Text Document.
		$LOB_REP_OUTPUT_TYPE_SPREADSHEET                                                 ; The Output Document when the Report is executed will be a Spreadsheet Document.

; Report Header/Footer Print Options.
Global Const _                                                                           ; com.sun.star.report.ReportPrintOption
		$LOB_REP_PAGE_PRINT_OPT_ALL_PAGES = 0, _                                         ; The page header/footer is printed on all pages.
		$LOB_REP_PAGE_PRINT_OPT_NOT_WITH_REP_HEADER = 1, _                               ; The page header/footer is not printed on the same page as the report header.
		$LOB_REP_PAGE_PRINT_OPT_NOT_WITH_REP_FOOTER = 2, _                               ; The page header/footer is not printed on the same page as the report footer.
		$LOB_REP_PAGE_PRINT_OPT_NOT_WITH_REP_HEADER_FOOTER = 3                           ; The page header/footer is not printed on the same page as the report header or footer.

; Report Section types.
Global Enum _
		$LOB_REP_SECTION_TYPE_DETAIL, _                                                  ; 0 A Report's Detail section.
		$LOB_REP_SECTION_TYPE_PAGE_FOOTER, _                                             ; 1 A Report's Page Footer section.
		$LOB_REP_SECTION_TYPE_PAGE_HEADER, _                                             ; 2 A Report's Page Header section.
		$LOB_REP_SECTION_TYPE_REPORT_FOOTER, _                                           ; 3 A Report's Report Footer section.
		$LOB_REP_SECTION_TYPE_REPORT_HEADER                                              ; 4 A Report's Report Header section.

; Result Set Cursor Movement Commands.
Global Enum _
		$LOB_RESULT_CURSOR_MOVE_BEFORE_FIRST, _                                          ; Moves before the first row.
		$LOB_RESULT_CURSOR_MOVE_FIRST, _                                                 ; Moves to the first row.
		$LOB_RESULT_CURSOR_MOVE_PREVIOUS, _                                              ; Moves back one row.
		$LOB_RESULT_CURSOR_MOVE_NEXT, _                                                  ; Moves forward one row.
		$LOB_RESULT_CURSOR_MOVE_LAST, _                                                  ; Moves to the last record.
		$LOB_RESULT_CURSOR_MOVE_AFTER_LAST, _                                            ; Moves after the last record.
		$LOB_RESULT_CURSOR_MOVE_ABSOLUTE, _                                              ; Moves to the row with the given row number.
		$LOB_RESULT_CURSOR_MOVE_RELATIVE                                                 ; Moves backwards or forwards by the given amount: forwards for a positive value, and backwards for a negative value.

; Result Set Cursor Queries.
Global Enum _
		$LOB_RESULT_CURSOR_QUERY_IS_BEFORE_FIRST, _                                      ; Is the cursor before the first record. This is the case if it has not yet been reset after entry.
		$LOB_RESULT_CURSOR_QUERY_IS_FIRST, _                                             ; Is the cursor on the first entry.
		$LOB_RESULT_CURSOR_QUERY_IS_LAST, _                                              ; Is the cursor on the last entry.
		$LOB_RESULT_CURSOR_QUERY_IS_AFTER_LAST, _                                        ; Is the cursor after the last row when it is moved on with next.
		$LOB_RESULT_CURSOR_QUERY_GET_ROW                                                 ; Retrieve the current row number.

; Column Nullability Constants.
Global Const _                                                                           ; com.sun.star.sdbc.ColumnValue Constant Group
		$LOB_RESULT_METADATA_COLUMN_NOT_NULLABLE = 0, _                                  ; The column does not allow NULL values.
		$LOB_RESULT_METADATA_COLUMN_NULLABLE = 1, _                                      ; The column does allow NULL values.
		$LOB_RESULT_METADATA_COLUMN_UNKNOWN_NULLABLE = 2                                 ; The nullability of the column is unknown.

; Column Metadata Query
Global Enum _
		$LOB_RESULT_METADATA_QUERY_GET_CATALOG_NAME, _                                   ; Gets a column's table's catalog name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_DECIMAL_PLACE, _                                  ; Gets a column's number of digits to right of the decimal point. Returns an Integer.
		$LOB_RESULT_METADATA_QUERY_GET_DISP_SIZE, _                                      ; Gets the column's normal max width in chars. Returns an Integer.
		$LOB_RESULT_METADATA_QUERY_GET_LABEL, _                                          ; Gets the suggested column title for use in printouts and displays. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_LENGTH, _                                         ; Gets a column's number of decimal digits. Returns an Integer.
		$LOB_RESULT_METADATA_QUERY_GET_NAME, _                                           ; Gets a column's name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_SCHEMA_NAME, _                                    ; Gets a column's table's schema. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_TABLE_NAME, _                                     ; Gets a column's table name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_TYPE, _                                           ; Gets the column's SQL type. Returns an Integer. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
		$LOB_RESULT_METADATA_QUERY_GET_TYPE_NAME, _                                      ; Gets the column's database-specific type name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_IS_AUTO_VALUE, _                                      ; Query whether the column is automatically numbered, thus read-only (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_CASE_SENSITIVE, _                                  ; Query whether a column's case matters (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_CURRENCY, _                                        ; Query whether the column is a cash value (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_NULLABLE, _                                        ; Query the nullability of values in the designated column. Returns an Integer. See Constants, $LOB_RESULT_METADATA_COLUMN_* as defined in LibreOfficeBase_Constants.au3.
		$LOB_RESULT_METADATA_QUERY_IS_READ_ONLY, _                                       ; Query whether a column is definitely not writable (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_SEARCHABLE, _                                      ; Query whether the column can be used in a where clause (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_SIGNED, _                                          ; Query whether values in the column are signed numbers (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_WRITABLE, _                                        ; Query whether it is possible for a write on the column to succeed (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_WRITABLE_DEFINITE                                  ; Query whether a write on the column will definitely succeed (True if so). Returns a Boolean.

; Result Set Row Modification Commands.
Global Enum _
		$LOB_RESULT_ROW_MOD_NULL, _                                                      ; Sets the column content to NULL
		$LOB_RESULT_ROW_MOD_BOOL, _                                                      ; Changes the content of specified column to the called logical value.
		$LOB_RESULT_ROW_MOD_BYTE, _                                                      ; Changes the content of specified column to the called byte.
		$LOB_RESULT_ROW_MOD_SHORT, _                                                     ; Changes the content of specified column to the called integer.
		$LOB_RESULT_ROW_MOD_INT, _                                                       ; Changes the content of specified column to the called integer.
		$LOB_RESULT_ROW_MOD_LONG, _                                                      ; Changes the content of specified column to the called integer.
		$LOB_RESULT_ROW_MOD_FLOAT, _                                                     ; Changes the content of specified column to the called decimal number.
		$LOB_RESULT_ROW_MOD_DOUBLE, _                                                    ; Changes the content of specified column to the called decimal number.
		$LOB_RESULT_ROW_MOD_STRING, _                                                    ; Changes the content of specified column to the called string.
		$LOB_RESULT_ROW_MOD_BYTES, _                                                     ; Changes the content of specified column to the called byte array.
		$LOB_RESULT_ROW_MOD_DATE, _                                                      ; Changes the content of specified column to the called date.
		$LOB_RESULT_ROW_MOD_TIME, _                                                      ; Changes the content of specified column to the called time.
		$LOB_RESULT_ROW_MOD_TIMESTAMP                                                    ; Changes the content of specified column to the called Date and Time (Timestamp).

; Result Set Queries.
Global Enum _
		$LOB_RESULT_ROW_QUERY_IS_ROW_INSERTED, _                                         ; Indicates if this is a new row.
		$LOB_RESULT_ROW_QUERY_IS_ROW_UPDATED, _                                          ; Indicates if the current row has been altered.
		$LOB_RESULT_ROW_QUERY_IS_ROW_DELETED                                             ; Indicates if the current row has been deleted.

; Result Set Row Read Commands.
Global Enum _
		$LOB_RESULT_ROW_READ_STRING, _                                                   ; Returns the content of the column as a character string.
		$LOB_RESULT_ROW_READ_BOOL, _                                                     ; Returns the content of the column as a boolean value.
		$LOB_RESULT_ROW_READ_BYTE, _                                                     ; Returns the content of the column as a single byte.
		$LOB_RESULT_ROW_READ_SHORT, _                                                    ; Returns the content of the column as an integer.
		$LOB_RESULT_ROW_READ_INT, _                                                      ; Returns the content of the column as an integer.
		$LOB_RESULT_ROW_READ_LONG, _                                                     ; Returns the content of the column as an integer.
		$LOB_RESULT_ROW_READ_FLOAT, _                                                    ; Returns the content of the column as a single precision decimal number.
		$LOB_RESULT_ROW_READ_DOUBLE, _                                                   ; Returns the content of the column as a double precision decimal number.
		$LOB_RESULT_ROW_READ_BYTES, _                                                    ; Returns the content of the column as an array of single bytes.
		$LOB_RESULT_ROW_READ_DATE, _                                                     ; Returns the content of the column as a date.
		$LOB_RESULT_ROW_READ_TIME, _                                                     ; Returns the content of the column as a time value.
		$LOB_RESULT_ROW_READ_TIMESTAMP, _                                                ; Returns the content of the column as a timestamp (date and time).
		$LOB_RESULT_ROW_READ_WAS_NULL                                                    ; Indicates if the value of the most recently read column was NULL.

; Result Set Row Update Commands.
Global Enum _
		$LOB_RESULT_ROW_UPDATE_INSERT, _                                                 ; Saves a new row.
		$LOB_RESULT_ROW_UPDATE_UPDATE, _                                                 ; Confirms alteration of the current row.
		$LOB_RESULT_ROW_UPDATE_DELETE, _                                                 ; Deletes the current row.
		$LOB_RESULT_ROW_UPDATE_CANCEL_UPDATE, _                                          ; Reverses changes in the current row.
		$LOB_RESULT_ROW_UPDATE_MOVE_TO_INSERT, _                                         ; Moves the cursor into a row corresponding to a new record.
		$LOB_RESULT_ROW_UPDATE_MOVE_TO_CURRENT                                           ; After the entry of a new record, returns the cursor to its previous position.

; Database Result Set Type Constants
Global Const _                                                                           ; com.sun.star.sdbc.ResultSetType Constant Group
		$LOB_RESULT_TYPE_FORWARD_ONLY = 1003, _                                          ; The Result Set cursor may move only forward.
		$LOB_RESULT_TYPE_SCROLL_INSENSITIVE = 1004, _                                    ; The Result Set is scrollable but generally not sensitive to changes made by others.
		$LOB_RESULT_TYPE_SCROLL_SENSITIVE = 1005                                         ; The Result Set is scrollable and generally sensitive to changes made by others.

; Strikeout
Global Const _                                                                           ; com.sun.star.awt.FontStrikeout
		$LOB_STRIKEOUT_NONE = 0, _                                                       ; No strike out.
		$LOB_STRIKEOUT_SINGLE = 1, _                                                     ; Strike out the characters with a single line.
		$LOB_STRIKEOUT_DOUBLE = 2, _                                                     ; Strike out the characters with a double line.
		$LOB_STRIKEOUT_DONT_KNOW = 3, _                                                  ; The strikeout mode is not specified. For Read Only.
		$LOB_STRIKEOUT_BOLD = 4, _                                                       ; Strike out the characters with a bold line.
		$LOB_STRIKEOUT_SLASH = 5, _                                                      ; Strike out the characters with slashes.
		$LOB_STRIKEOUT_X = 6                                                             ; Strike out the characters with X's.

; Database SubComponent type.
Global Const _                                                                           ; com.sun.star.sdb.application.DatabaseObject
		$LOB_SUB_COMP_TYPE_TABLE = 0, _                                                  ; A Database Table.
		$LOB_SUB_COMP_TYPE_QUERY = 1, _                                                  ; A Database Query.
		$LOB_SUB_COMP_TYPE_FORM = 2, _                                                   ; A Database Form.
		$LOB_SUB_COMP_TYPE_REPORT = 3                                                    ; A Database Report.

; Text Horizontal Alignment
Global Const _                                                                           ; com.sun.star.style.ParagraphAdjust
		$LOB_TXT_ALIGN_HORI_LEFT = 0, _                                                  ; Text is adjusted to the left border.
		$LOB_TXT_ALIGN_HORI_RIGHT = 1, _                                                 ; Text is adjusted to the right border.
		$LOB_TXT_ALIGN_HORI_BLOCK = 2, _                                                 ; Text is adjusted to both borders / stretched, except for last line.
		$LOB_TXT_ALIGN_HORI_CENTER = 3, _                                                ; Text is adjusted to the center.
		$LOB_TXT_ALIGN_HORI_STRETCH = 4                                                  ; Text is adjusted to both borders / stretched, including last line.

; Underline/Overline
Global Const _                                                                           ; com.sun.star.awt.FontUnderline
		$LOB_UNDERLINE_NONE = 0, _                                                       ; No Underline or Overline style.
		$LOB_UNDERLINE_SINGLE = 1, _                                                     ; Single line Underline/Overline style.
		$LOB_UNDERLINE_DOUBLE = 2, _                                                     ; Double line Underline/Overline style.
		$LOB_UNDERLINE_DOTTED = 3, _                                                     ; Dotted line Underline/Overline style.
		$LOB_UNDERLINE_DONT_KNOW = 4, _                                                  ; Unknown Underline/Overline style, for read only.
		$LOB_UNDERLINE_DASH = 5, _                                                       ; Dashed line Underline/Overline style.
		$LOB_UNDERLINE_LONG_DASH = 6, _                                                  ; Long Dashed line Underline/Overline style.
		$LOB_UNDERLINE_DASH_DOT = 7, _                                                   ; Dash Dot line Underline/Overline style.
		$LOB_UNDERLINE_DASH_DOT_DOT = 8, _                                               ; Dash Dot Dot line Underline/Overline style.
		$LOB_UNDERLINE_SML_WAVE = 9, _                                                   ; Small Wave line Underline/Overline style.
		$LOB_UNDERLINE_WAVE = 10, _                                                      ; Wave line Underline/Overline style.
		$LOB_UNDERLINE_DBL_WAVE = 11, _                                                  ; Double Wave line Underline/Overline style.
		$LOB_UNDERLINE_BOLD = 12, _                                                      ; Bold line Underline/Overline style.
		$LOB_UNDERLINE_BOLD_DOTTED = 13, _                                               ; Bold Dotted line Underline/Overline style.
		$LOB_UNDERLINE_BOLD_DASH = 14, _                                                 ; Bold Dashed line Underline/Overline style.
		$LOB_UNDERLINE_BOLD_LONG_DASH = 15, _                                            ; Bold Long Dash line Underline/Overline style.
		$LOB_UNDERLINE_BOLD_DASH_DOT = 16, _                                             ; Bold Dash Dot line Underline/Overline style.
		$LOB_UNDERLINE_BOLD_DASH_DOT_DOT = 17, _                                         ; Bold Dash Dot Dot line Underline/Overline style.
		$LOB_UNDERLINE_BOLD_WAVE = 18                                                    ; Bold Wave line Underline/Overline style.

; Weight/Bold
Global Const _                                                                           ; com.sun.star.awt.FontWeight
		$LOB_WEIGHT_DONT_KNOW = 0, _                                                     ; The font weight is not specified/unknown. For Read Only.
		$LOB_WEIGHT_THIN = 50, _                                                         ; A 50% (Thin) font weight.
		$LOB_WEIGHT_ULTRA_LIGHT = 60, _                                                  ; A 60% (Ultra Light) font weight.
		$LOB_WEIGHT_LIGHT = 75, _                                                        ; A 75% (Light) font weight.
		$LOB_WEIGHT_SEMI_LIGHT = 90, _                                                   ; A 90% (Semi-Light) font weight.
		$LOB_WEIGHT_NORMAL = 100, _                                                      ; A 100% (Normal) font weight.
		$LOB_WEIGHT_SEMI_BOLD = 110, _                                                   ; A 110% (Semi-Bold) font weight.
		$LOB_WEIGHT_BOLD = 150, _                                                        ; A 150% (Bold) font weight.
		$LOB_WEIGHT_ULTRA_BOLD = 175, _                                                  ; A 175% (Ultra-Bold) font weight.
		$LOB_WEIGHT_BLACK = 200                                                          ; A 200% (Black) font weight.
