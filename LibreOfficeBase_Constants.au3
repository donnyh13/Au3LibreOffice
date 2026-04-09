#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice Base Constants for the LibreOffice UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various functions in the LibreOffice UDF.
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
		$LOB_CHAR_CASEMAP_NONE = 0, _                                                    ; The case of the characters is unchanged.
		$LOB_CHAR_CASEMAP_UPPER = 1, _                                                   ; All characters are put in upper case.
		$LOB_CHAR_CASEMAP_LOWER = 2, _                                                   ; All characters are put in lower case.
		$LOB_CHAR_CASEMAP_TITLE = 3, _                                                   ; The first character of each word is put in upper case.
		$LOB_CHAR_CASEMAP_SM_CAPS = 4                                                    ; All characters are put in upper case, but with a smaller font height.

; Posture/Italic
Global Const _                                                                           ; com.sun.star.awt.FontSlant
		$LOB_CHAR_POSTURE_NONE = 0, _                                                    ; Specifies a font without slant.
		$LOB_CHAR_POSTURE_OBLIQUE = 1, _                                                 ; Specifies an oblique font (slant not designed into the font).
		$LOB_CHAR_POSTURE_ITALIC = 2, _                                                  ; Specifies an italic font (slant designed into the font).
		$LOB_CHAR_POSTURE_DONTKNOW = 3, _                                                ; Specifies a font with an unknown slant. For Read Only.
		$LOB_CHAR_POSTURE_REV_OBLIQUE = 4, _                                             ; Specifies a reverse oblique font (slant not designed into the font).
		$LOB_CHAR_POSTURE_REV_ITALIC = 5                                                 ; Specifies a reverse italic font (slant designed into the font).

; Relief
Global Const _                                                                           ; com.sun.star.text.FontRelief
		$LOB_CHAR_RELIEF_NONE = 0, _                                                     ; No relief is applied.
		$LOB_CHAR_RELIEF_EMBOSSED = 1, _                                                 ; The font relief is embossed.
		$LOB_CHAR_RELIEF_ENGRAVED = 2                                                    ; The font relief is engraved.

; Strikeout
Global Const _                                                                           ; com.sun.star.awt.FontStrikeout
		$LOB_CHAR_STRIKEOUT_NONE = 0, _                                                  ; No strike out.
		$LOB_CHAR_STRIKEOUT_SINGLE = 1, _                                                ; Strike out the characters with a single line.
		$LOB_CHAR_STRIKEOUT_DOUBLE = 2, _                                                ; Strike out the characters with a double line.
		$LOB_CHAR_STRIKEOUT_DONT_KNOW = 3, _                                             ; The strikeout mode is not specified. For Read Only.
		$LOB_CHAR_STRIKEOUT_BOLD = 4, _                                                  ; Strike out the characters with a bold line.
		$LOB_CHAR_STRIKEOUT_SLASH = 5, _                                                 ; Strike out the characters with slashes.
		$LOB_CHAR_STRIKEOUT_X = 6                                                        ; Strike out the characters with X's.

; Underline/Overline
Global Const _                                                                           ; com.sun.star.awt.FontUnderline
		$LOB_CHAR_UNDERLINE_NONE = 0, _                                                  ; No Underline or Overline style.
		$LOB_CHAR_UNDERLINE_SINGLE = 1, _                                                ; Single line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_DOUBLE = 2, _                                                ; Double line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_DOTTED = 3, _                                                ; Dotted line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_DONT_KNOW = 4, _                                             ; Unknown Underline/Overline style, for read only.
		$LOB_CHAR_UNDERLINE_DASH = 5, _                                                  ; Dashed line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_LONG_DASH = 6, _                                             ; Long Dashed line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_DASH_DOT = 7, _                                              ; Dash Dot line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_DASH_DOT_DOT = 8, _                                          ; Dash Dot Dot line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_SML_WAVE = 9, _                                              ; Small Wave line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_WAVE = 10, _                                                 ; Wave line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_DBL_WAVE = 11, _                                             ; Double Wave line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_BOLD = 12, _                                                 ; Bold line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_BOLD_DOTTED = 13, _                                          ; Bold Dotted line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_BOLD_DASH = 14, _                                            ; Bold Dashed line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_BOLD_LONG_DASH = 15, _                                       ; Bold Long Dash line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_BOLD_DASH_DOT = 16, _                                        ; Bold Dash Dot line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_BOLD_DASH_DOT_DOT = 17, _                                    ; Bold Dash Dot Dot line Underline/Overline style.
		$LOB_CHAR_UNDERLINE_BOLD_WAVE = 18                                               ; Bold Wave line Underline/Overline style.

; Weight/Bold
Global Const _                                                                           ; com.sun.star.awt.FontWeight
		$LOB_CHAR_WEIGHT_DONT_KNOW = 0, _                                                ; The font weight is not specified/unknown. For Read Only.
		$LOB_CHAR_WEIGHT_THIN = 50, _                                                    ; A 50% (Thin) font weight.
		$LOB_CHAR_WEIGHT_ULTRA_LIGHT = 60, _                                             ; A 60% (Ultra Light) font weight.
		$LOB_CHAR_WEIGHT_LIGHT = 75, _                                                   ; A 75% (Light) font weight.
		$LOB_CHAR_WEIGHT_SEMI_LIGHT = 90, _                                              ; A 90% (Semi-Light) font weight.
		$LOB_CHAR_WEIGHT_NORMAL = 100, _                                                 ; A 100% (Normal) font weight.
		$LOB_CHAR_WEIGHT_SEMI_BOLD = 110, _                                              ; A 110% (Semi-Bold) font weight.
		$LOB_CHAR_WEIGHT_BOLD = 150, _                                                   ; A 150% (Bold) font weight.
		$LOB_CHAR_WEIGHT_ULTRA_BOLD = 175, _                                             ; A 175% (Ultra-Bold) font weight.
		$LOB_CHAR_WEIGHT_BLACK = 200                                                     ; A 200% (Black) font weight.

; Table Column Text Alignment
Global Const _                                                                           ; com.sun.star.sdb.Align
		$LOB_COL_TXT_ALIGN_LEFT = 0, _                                                   ; The Column's Text is aligned to the Left.
		$LOB_COL_TXT_ALIGN_CENTER = 1, _                                                 ; The Column's Text is aligned in the center.
		$LOB_COL_TXT_ALIGN_RIGHT = 2                                                     ; The Column's Text is aligned to the Right.

; Prepared Statement Input Type Commands.
Global Enum _
		$LOB_DATA_SET_TYPE_NULL, _                                                       ; 0 Sets the content of the column to NULL.
		$LOB_DATA_SET_TYPE_BOOL, _                                                       ; 1 Puts the given logical value into the SQL command.
		$LOB_DATA_SET_TYPE_BYTE, _                                                       ; 2 Puts the given byte into the SQL command.
		$LOB_DATA_SET_TYPE_SHORT, _                                                      ; 3 Puts the given integer into the SQL command.
		$LOB_DATA_SET_TYPE_INT, _                                                        ; 4 Puts the given integer into the SQL command.
		$LOB_DATA_SET_TYPE_LONG, _                                                       ; 5 Puts the given integer into the SQL command.
		$LOB_DATA_SET_TYPE_FLOAT, _                                                      ; 6 Puts the given decimal number into the SQL command.
		$LOB_DATA_SET_TYPE_DOUBLE, _                                                     ; 7 Puts the given decimal number into the SQL command.
		$LOB_DATA_SET_TYPE_STRING, _                                                     ; 8 Puts the given character string into the SQL command.
		$LOB_DATA_SET_TYPE_BYTES, _                                                      ; 9 Puts the given byte array into the SQL command.
		$LOB_DATA_SET_TYPE_DATE, _                                                       ; 10 Puts the given date into the SQL command.
		$LOB_DATA_SET_TYPE_TIME, _                                                       ; 11 Puts the given time into the SQL command.
		$LOB_DATA_SET_TYPE_TIMESTAMP, _                                                  ; 12 Puts the given timestamp into the SQL command.
		$LOB_DATA_SET_TYPE_CLOB, _                                                       ; 13 Puts the given CLOB (Character Large Object) into the SQL command.
		$LOB_DATA_SET_TYPE_BLOB, _                                                       ; 14 Puts the given BLOB (Binary Large Object) into the SQL command.
		$LOB_DATA_SET_TYPE_ARRAY, _                                                      ; 15 Puts the given Array into the SQL command.
		$LOB_DATA_SET_TYPE_OBJECT                                                        ; 16 Puts the given Object into the SQL command.

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
		$LOB_DBASE_META_ALL_PROCEDURES_ARE_CALLABLE, _                                   ; 0 Returns a Boolean. Can all the procedures returned by getProcedures be called by the current user? TRUE if the user is allowed to call all procedures returned by getProcedures otherwise FALSE.
		$LOB_DBASE_META_ALL_TABLES_ARE_SELECTABLE, _                                     ; 1 Returns a Boolean. Can all the tables returned by getTable be SELECTed by the current user? True if so.
		$LOB_DBASE_META_DATA_DEFINITION_CAUSES_TRANSACTION_COMMIT, _                     ; 2 Returns a Boolean. Does a data definition statement within a transaction force the transaction to commit? True if so.
		$LOB_DBASE_META_DATA_DEFINITION_IGNORED_IN_TRANSACTIONS, _                       ; 3 Returns a Boolean. Is a data definition statement within a transaction ignored? True if so.
		$LOB_DBASE_META_DELETES_ARE_DETECTED, _                                          ; 4 Returns a Boolean. Indicates whether or not a visible row delete can be detected by calling $LOB_RESULT_ROW_QUERY_IS_ROW_DELETED. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_DOES_MAX_ROW_SIZE_INCLUDE_BLOBS, _                               ; 5 Did $LOB_DBASE_META_GET_MAX_ROW_SIZE include LONGVARCHAR and LONGVARBINARY blobs? Returns a Boolean. True if so.
		$LOB_DBASE_META_GET_BEST_ROW_ID, _                                               ; 6 Returns a Result Set, each row is a column description. Gets a description of a table's optimal set of columns that uniquely identifies a row. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_CATALOG_SEPARATOR, _                                         ; 7 Returns a String. Return the separator between catalog and table name.
		$LOB_DBASE_META_GET_CATALOG_TERM, _                                              ; 8 Returns a String. Return the database vendor's preferred term for "catalog".
		$LOB_DBASE_META_GET_CATALOGS, _                                                  ; 9 Returns a Result Set, each row has a single String column that is a catalog name. Gets the catalog names available in this database. The catalog column is: TABLE_CAT string => catalog name.
		$LOB_DBASE_META_GET_COLS, _                                                      ; 10 Returns a Result Set, each row is a column privilege description. Gets a description of table columns available in the specified catalog. Only column descriptions matching the catalog, schema, table and column name criteria are returned. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_COL_PRIVILEGES, _                                            ; 11 Returns a Result Set, each row is a column description. Gets a description of the access rights for a table's columns. Only privileges matching the column name criteria are returned. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_CROSS_REF, _                                                 ; 12 Returns a Result Set, each row is a foreign key column description. Gets a description of the foreign key columns in the foreign key table that reference the primary key columns of the primary key table (describe how one table imports another's key.) This should normally return a single foreign key/primary key pair (most tables only import a foreign key from a table once.). See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_DATABASE_PRODUCT_NAME, _                                     ; 13 Returns a String. Returns the name of the database product.
		$LOB_DBASE_META_GET_DATABASE_PRODUCT_VERSION, _                                  ; 14 Returns a String. Returns the version of the database product.
		$LOB_DBASE_META_GET_DEFAULT_TRANSACTION_ISOLATION, _                             ; 15 Returns an Integer. Return the database default transaction isolation level. See $LOB_DBASE_TRANSACTION_ISOLATION_* Constants.
		$LOB_DBASE_META_GET_DRIVER_MAJOR_VERSION, _                                      ; 16 Returns an Integer. Returns the SDBC driver major version number.
		$LOB_DBASE_META_GET_DRIVER_MINOR_VERSION, _                                      ; 17 Returns an Integer. Returns the SDBC driver minor version number.
		$LOB_DBASE_META_GET_DRIVER_NAME, _                                               ; 18 Returns a String. Returns the name of the SDBC driver.
		$LOB_DBASE_META_GET_DRIVER_VERSION, _                                            ; 19 Returns an Integer. Returns the version number of the SDBC driver.
		$LOB_DBASE_META_GET_EXPORTED_KEYS, _                                             ; 20 Returns a Result Set, each row is a foreign key column description. Gets a description of the foreign key columns that reference a table's primary key columns (the foreign keys exported by a table). See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_EXTRA_NAME_CHARS, _                                          ; 21 Returns a String. Gets all the "extra" characters that can be used in unquoted identifier names (those beyond a-z, A-Z, 0-9 and _).
		$LOB_DBASE_META_GET_IDENTIFIER_QUOTE_STRING, _                                   ; 22 Returns a String. What's the string used to quote SQL identifiers? This returns a space " " if identifier quoting is not supported.
		$LOB_DBASE_META_GET_IMPORTED_KEYS, _                                             ; 23 Returns a Result Set, each row is a primary key column description. Gets a description of the primary key columns that are referenced by a table's foreign key columns (the primary keys imported by a table). See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_INDEX_INFO, _                                                ; 24 Returns a Result Set, each row is an index column description. Gets a description of a table's indices and statistics. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_MAX_BINARY_LITERAL_LEN, _                                    ; 25 Returns an Integer. Return the maximal number of hex characters in an inline binary literal.
		$LOB_DBASE_META_GET_MAX_CATALOG_NAME_LEN, _                                      ; 26 Returns an Integer. Return the maximum length of a catalog name.
		$LOB_DBASE_META_GET_MAX_CHAR_LITERAL_LEN, _                                      ; 27 Returns an Integer. Return the max length for a character literal.
		$LOB_DBASE_META_GET_MAX_COL_NAME_LEN, _                                          ; 28 Returns an Integer. Return the limit on column name length.
		$LOB_DBASE_META_GET_MAX_COLS_IN_GROUP_BY, _                                      ; 29 Returns an Integer. Return the maximum number of columns in a "GROUP BY" clause.
		$LOB_DBASE_META_GET_MAX_COLS_IN_INDEX, _                                         ; 30 Returns an Integer. Return the maximum number of columns allowed in an index.
		$LOB_DBASE_META_GET_MAX_COLS_IN_ORDER_BY, _                                      ; 31 Returns an Integer. Return the maximum number of columns in an "ORDER BY" clause.
		$LOB_DBASE_META_GET_MAX_COLS_IN_SEL, _                                           ; 32 Returns an Integer. Return the maximum number of columns in a "SELECT" list.
		$LOB_DBASE_META_GET_MAX_COLS_IN_TABLE, _                                         ; 33 Returns an Integer. Return the maximum number of columns in a table.
		$LOB_DBASE_META_GET_MAX_CONNECTIONS, _                                           ; 34 Returns an Integer. Return the number of active connections at a time to this database.
		$LOB_DBASE_META_GET_MAX_CURSOR_NAME_LEN, _                                       ; 35 Returns an Integer. Return the maximum cursor name length.
		$LOB_DBASE_META_GET_MAX_INDEX_LEN, _                                             ; 36 Returns an Integer. Return the maximum length of an index (in bytes).
		$LOB_DBASE_META_GET_MAX_PROCEDURE_NAME_LEN, _                                    ; 37 Returns an Integer. Return the maximum length of a procedure name.
		$LOB_DBASE_META_GET_MAX_ROW_SIZE, _                                              ; 38 Returns an Integer. Return the maximum length of a single row.
		$LOB_DBASE_META_GET_MAX_SCHEMA_NAME_LEN, _                                       ; 39 Returns an Integer. Return the maximum length allowed for a schema name.
		$LOB_DBASE_META_GET_MAX_STATEMENT_LEN, _                                         ; 40 Returns an Integer. Return the maximum length of a SQL statement.
		$LOB_DBASE_META_GET_MAX_STATEMENTS, _                                            ; 41 Returns an Integer. Return the maximal number of open active statements at one time to this database.
		$LOB_DBASE_META_GET_MAX_TABLE_NAME_LEN, _                                        ; 42 Returns an Integer. Return the maximum length of a table name.
		$LOB_DBASE_META_GET_MAX_TABLES_IN_SEL, _                                         ; 43 Returns an Integer. Return the maximum number of tables in a SELECT statement.
		$LOB_DBASE_META_GET_MAX_USER_NAME_LEN, _                                         ; 44 Returns an Integer. Return the maximum length of a user name.
		$LOB_DBASE_META_GET_NUMERIC_FUNCS, _                                             ; 45 Returns a String. Gets a comma-separated list of math functions. These are the X/Open CLI math function names used in the SDBC function escape clause.
		$LOB_DBASE_META_GET_PRIMARY_KEY, _                                               ; 46 Returns a Result Set, each row is a primary key column description. Gets a description of a table's primary key columns. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_PROCEDURE_COLS, _                                            ; 47 Returns a Result Set, each row describes a stored procedure parameter or column. Gets a description of a catalog's stored procedure parameters and result columns. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_PROCEDURE_TERM, _                                            ; 48 Returns a String. Return the database vendor's preferred term for "procedure".
		$LOB_DBASE_META_GET_PROCEDURES, _                                                ; 49 Returns a Result Set, each row is a procedure description. Gets a description of the stored procedures available in a catalog. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_SCHEMA_TERM, _                                               ; 50 Returns a String. Return the database vendor's preferred term for "schema".
		$LOB_DBASE_META_GET_SCHEMAS, _                                                   ; 51 Returns a Result Set, each row has a single String column that is a schema name. Gets the schema names available in this database. The schema column is: TABLE_SCHEM string => schema name.
		$LOB_DBASE_META_GET_SEARCH_STRING_ESCAPE, _                                      ; 52 Returns a String. Gets the string that can be used to escape wildcard characters. This is the string that can be used to escape "_" or "%" in the string pattern style catalog search parameters.
		$LOB_DBASE_META_GET_SQL_KEYWORDS, _                                              ; 53 Returns a String. Gets a comma-separated list of all a database's SQL keywords that are NOT also SQL92 keywords.
		$LOB_DBASE_META_GET_STRING_FUNCS, _                                              ; 54 Returns a String. Gets a comma-separated list of string functions. These are the X/Open CLI string function names used in the SDBC function escape clause.
		$LOB_DBASE_META_GET_SYSTEM_FUNCS, _                                              ; 55 Returns a String. Gets a comma-separated list of system functions. These are the X/Open CLI system function names used in the SDBC function escape clause.
		$LOB_DBASE_META_GET_TABLE_PRIVILEGES, _                                          ; 56 Returns a Result Set, each row is a table privilege description. Gets a description of the access rights for each table available in a catalog. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_TABLE_TYPES, _                                               ; 57 Returns a Result Set, each row has a single String column that is a table type. Gets the table types available in this database.  See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_TABLES, _                                                    ; 58 Returns a Result Set, each row is a table description.  Gets a description of tables available in a catalog. See _LOBase_DatabaseMetaDataQuery for Parameters.See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_TIME_DATE_FUNCS, _                                           ; 59 Returns a String. Gets a comma-separated list of time and date functions.
		$LOB_DBASE_META_GET_TYPE_INFO, _                                                 ; 60 Returns a Result Set, each row is a SQL type description. Gets a description of all the standard SQL types supported by this database. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_UDTS, _                                                      ; 61 Returns a Result Set, each row is a type description. Gets a description of the user-defined types defined in a particular schema. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_VERSION_COLS, _                                              ; 62 Returns a Result Set, each row is a column description. Gets a description of a table's columns that are automatically updated when any value in a row is updated. See _LOBase_DatabaseMetaDataQuery for Parameters. See L.O. SDL API "XDatabaseMetaData Interface Reference" for more details.
		$LOB_DBASE_META_GET_URL, _                                                       ; 63 Returns a String. Returns the URL for the database connection.
		$LOB_DBASE_META_GET_USERNAME, _                                                  ; 64 Returns a String. Returns the user name from this database connection.
		$LOB_DBASE_META_INSERTS_ARE_DETECTED, _                                          ; 65 Returns a Boolean. Indicates whether or not a visible row insert can be detected by calling $LOB_RESULT_ROW_QUERY_IS_ROW_INSERTED. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_IS_CATALOG_AT_START, _                                           ; 66 Returns a Boolean. Does a catalog appear at the start of a qualified table name? (Otherwise it appears at the end). True if so.
		$LOB_DBASE_META_IS_READ_ONLY, _                                                  ; 67 Returns a Boolean. Checks if the database in read-only mode. True if so.
		$LOB_DBASE_META_NULL_PLUS_NON_NULL_IS_NULL, _                                    ; 68 Returns a Boolean. Are concatenations between NULL and non-NULL values NULL? True if so.
		$LOB_DBASE_META_NULLS_ARE_SORTED_AT_END, _                                       ; 69 Returns a Boolean. Are NULL values sorted at the end, regardless of sort order? True if so.
		$LOB_DBASE_META_NULLS_ARE_SORTED_AT_START, _                                     ; 70 Returns a Boolean. Are NULL values sorted at the start regardless of sort order? True if so.
		$LOB_DBASE_META_NULLS_ARE_SORTED_HIGH, _                                         ; 71 Returns a Boolean. Are NULL values sorted high? True if so.
		$LOB_DBASE_META_NULLS_ARE_SORTED_LOW, _                                          ; 72 Returns a Boolean. Are NULL values sorted low? True if so.
		$LOB_DBASE_META_OTHERS_DELETES_ARE_VISIBLE, _                                    ; 73 Returns a Boolean. Indicates whether deletes made by others are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_OTHERS_INSERTS_ARE_VISIBLE, _                                    ; 74 Returns a Boolean. Indicates whether inserts made by others are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_OTHERS_UPDATES_ARE_VISIBLE, _                                    ; 75 Returns a Boolean. Indicates whether updates made by others are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_OWN_DELETES_ARE_VISIBLE, _                                       ; 76 Returns a Boolean. Indicates whether a result set's own deletes are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_OWN_INSERTS_ARE_VISIBLE, _                                       ; 77 Returns a Boolean. Indicates whether a result set's own inserts are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_OWN_UPDATES_ARE_VISIBLE, _                                       ; 78 Returns a Boolean. Indicates whether a result set's own updates are visible. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_STORES_LOWER_CASE_IDS, _                                         ; 79 Returns a Boolean. Does the database treat mixed case unquoted SQL identifiers as case insensitive and store them in lower case? True if so.
		$LOB_DBASE_META_STORES_MIXED_CASE_IDS, _                                         ; 80 Returns a Boolean. Does the database treat mixed case unquoted SQL identifiers as case insensitive and store them in mixed case? True if so.
		$LOB_DBASE_META_STORES_UPPER_CASE_IDS, _                                         ; 81 Returns a Boolean. Does the database treat mixed case unquoted SQL identifiers as case insensitive and store them in upper case? True if so.
		$LOB_DBASE_META_STORES_LOWER_CASE_QUOTED_IDS, _                                  ; 82 Returns a Boolean. Does the database treat mixed case quoted SQL identifiers as case insensitive and store them in lower case? True if so.
		$LOB_DBASE_META_STORES_MIXED_CASE_QUOTED_IDS, _                                  ; 83 Returns a Boolean. Does the database treat mixed case quoted SQL identifiers as case insensitive and store them in mixed case? True if so.
		$LOB_DBASE_META_STORES_UPPER_CASE_QUOTED_IDS, _                                  ; 84 Returns a Boolean. Does the database treat mixed case quoted SQL identifiers as case insensitive and store them in upper case? True if so.
		$LOB_DBASE_META_SUPPORTS_ALTER_TABLE_WITH_ADD_COL, _                             ; 85 Returns a Boolean. Support the Database "ALTER TABLE" with add column? True if so.
		$LOB_DBASE_META_SUPPORTS_ALTER_TABLE_WITH_DROP_COL, _                            ; 86 Returns a Boolean. Support the Database "ALTER TABLE" with drop column? True if so.
		$LOB_DBASE_META_SUPPORTS_ANSI92_ENTRY_LEVEL_SQL, _                               ; 87 Returns a Boolean. TRUE, if the database supports ANSI92 entry level SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_ANSI92_FULL_SQL, _                                      ; 88 Returns a Boolean. TRUE, if the database supports ANSI92 full SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_ANSI92_INTERMEDIATE_SQL, _                              ; 89 Returns a Boolean. TRUE, if the database supports ANSI92 intermediate SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_BATCH_UPDATES, _                                        ; 90 Returns a Boolean. Indicates whether the driver supports batch updates. True if so.
		$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_DATA_MANIPULATION, _                        ; 91 Returns a Boolean. Can a catalog name be used in a data manipulation statement? True if so.
		$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_INDEX_DEFINITIONS, _                        ; 92 Returns a Boolean. Can a catalog name be used in an index definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_PRIVILEGE_DEFINITIONS, _                    ; 93 Returns a Boolean. Can a catalog name be used in a privilege definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_PROCEDURE_CALLS, _                          ; 94 Returns a Boolean. Can a catalog name be used in a procedure call statement? True if so.
		$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_TABLE_DEFINITION, _                         ; 95 Returns a Boolean. Can a catalog name be used in a table definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_COL_ALIASING, _                                         ; 96 Returns a Boolean. Support the Database column aliasing? The SQL AS clause can be used to provide names for computed columns or to provide alias names for columns as required. True if so.
		$LOB_DBASE_META_SUPPORTS_CONVERT, _                                              ; 97 Returns a Boolean. TRUE , if the Database supports the CONVERT between the given SQL types otherwise FALSE. See _LOBase_DatabaseMetaDataQuery for Parameters.
		$LOB_DBASE_META_SUPPORTS_CORE_SQL_GRAMMAR, _                                     ; 98 Returns a Boolean. TRUE, if the database supports ODBC Core SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_CORRELATED_SUBQUERIES, _                                ; 99 Returns a Boolean. Are correlated subqueries supported? True if so.
		$LOB_DBASE_META_SUPPORTS_DATA_DEFINITION_AND_DATA_MANIPULATION_TRANSACTIONS, _   ; 100 Returns a Boolean. Support the Database both data definition and data manipulation statements within a transaction? True if so.
		$LOB_DBASE_META_SUPPORTS_DATA_MANIPULATION_TRANSACTIONS_ONLY, _                  ; 101 Returns a Boolean. Are only data manipulation statements within a transaction supported? True if so.
		$LOB_DBASE_META_SUPPORTS_DIFF_TABLE_CORRELATION_NAMES, _                         ; 102 Returns a Boolean. If table correlation names are supported, are they restricted to be different from the names of the tables? True if so.
		$LOB_DBASE_META_SUPPORTS_EXPRESSIONS_IN_ORDER_BY, _                              ; 103 Returns a Boolean. Are expressions in "ORDER BY" lists supported? True if so.
		$LOB_DBASE_META_SUPPORTS_EXTENDED_SQL_GRAMMAR, _                                 ; 104 Returns a Boolean. TRUE, if the database supports ODBC Extended SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_FULL_OUTER_JOINS, _                                     ; 105 Returns a Boolean. TRUE, if full nested outer joins are supported, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_GROUP_BY, _                                             ; 106 Returns a Boolean. Is some form of "GROUP BY" clause supported? True if so.
		$LOB_DBASE_META_SUPPORTS_GROUP_BY_BEYOND_SELECT, _                               ; 107 Returns a Boolean. Can a "GROUP BY" clause add columns not in the SELECT provided it specifies all the columns in the SELECT? True if so.
		$LOB_DBASE_META_SUPPORTS_GROUP_BY_UNRELATED, _                                   ; 108 Returns a Boolean. Can a "GROUP BY" clause use columns not in the SELECT? True if so.
		$LOB_DBASE_META_SUPPORTS_INTEGRITY_ENHANCMENT_FACILITY, _                        ; 109 Returns a Boolean. TRUE, if the Database supports SQL Integrity Enhancement Facility, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_LIKE_ESCAPE_CLAUSE, _                                   ; 110 Returns a Boolean. Is the escape character in "LIKE" clauses supported? True if so.
		$LOB_DBASE_META_SUPPORTS_LIMITED_OUTER_JOINS, _                                  ; 111 Returns a Boolean. TRUE, if there is limited support for outer joins. (This will be TRUE if supportFullOuterJoins is TRUE.) FALSE is returned otherwise.
		$LOB_DBASE_META_SUPPORTS_MINIMUM_SQL_GRAMMAR, _                                  ; 112 Returns a Boolean. TRUE, if the database supports ODBC Minimum SQL grammar, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_MIXED_CASE_IDS, _                                       ; 113 Returns a Boolean. Use the database "mixed case unquoted SQL identifiers" case sensitive. True if so.
		$LOB_DBASE_META_SUPPORTS_MIXED_CASE_QUOTED_IDS, _                                ; 114 Returns a Boolean.  Does the database treat mixed case quoted SQL identifiers as case sensitive and as a result store them in mixed case?True if so.
		$LOB_DBASE_META_SUPPORTS_MULTIPLE_RESULT_SETS, _                                 ; 115 Returns a Boolean. Are multiple XResultSets from a single execute supported? True if so.
		$LOB_DBASE_META_SUPPORTS_MULTIPLE_TRANSACTIONS, _                                ; 116 Returns a Boolean. Can we have multiple transactions open at once (on different connections)? True if so.
		$LOB_DBASE_META_SUPPORTS_NON_NULLABLE_COLS, _                                    ; 117 Returns a Boolean. Can columns be defined as non-nullable? True if so.
		$LOB_DBASE_META_SUPPORTS_OPEN_CURSORS_ACROSS_COMMIT, _                           ; 118 Returns a Boolean. Can cursors remain open across commits? True if so.
		$LOB_DBASE_META_SUPPORTS_OPEN_CURSORS_ACROSS_ROLLBACK, _                         ; 119 Returns a Boolean. Can cursors remain open across rollbacks? True if so.
		$LOB_DBASE_META_SUPPORTS_OPEN_STATEMENTS_ACROSS_COMMIT, _                        ; 120 Returns a Boolean. Can statements remain open across commits? True if so.
		$LOB_DBASE_META_SUPPORTS_OPEN_STATEMENTS_ACROSS_ROLLBACK, _                      ; 121 Returns a Boolean. Can statements remain open across rollbacks? True if so.
		$LOB_DBASE_META_SUPPORTS_ORDER_BY_UNRELATED, _                                   ; 122 Returns a Boolean. Can an "ORDER BY" clause use columns not in the SELECT statement? True if so.
		$LOB_DBASE_META_SUPPORTS_OUTER_JOINS, _                                          ; 123 Returns a Boolean. TRUE, if some form of outer join is supported, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_POSITIONED_DELETE, _                                    ; 124 Returns a Boolean. Is positioned DELETE supported? True if so.
		$LOB_DBASE_META_SUPPORTS_POSITIONED_UPDATE, _                                    ; 125 Returns a Boolean. Is positioned UPDATE supported? True if so.
		$LOB_DBASE_META_SUPPORTS_RESULT_SET_CONCURRENCY, _                               ; 126 Returns a Boolean. Does the database support the concurrency type in combination with the given result set type? See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_SUPPORTS_RESULT_SET_TYPE, _                                      ; 127 Returns a Boolean. Does the database support the given result set type? See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_DATA_MANIPULATION, _                         ; 128 Returns a Boolean. Can a schema name be used in a data manipulation statement? True if so.
		$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_INDEX_DEFINITIONS, _                         ; 129 Returns a Boolean. Can a schema name be used in an index definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_PRIVILEGE_DEFINITIONS, _                     ; 130 Returns a Boolean. Can a schema name be used in a privilege definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_PROCEDURE_CALLS, _                           ; 131 Returns a Boolean. Can a schema name be used in a procedure call statement? True if so.
		$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_TABLE_DEFINITION, _                          ; 132 Returns a Boolean. Can a schema name be used in a table definition statement? True if so.
		$LOB_DBASE_META_SUPPORTS_SELECT_FOR_UPDATE, _                                    ; 133 Returns a Boolean. Is SELECT for UPDATE supported? True if so.
		$LOB_DBASE_META_SUPPORTS_STORED_PROCEDURES, _                                    ; 134 Returns a Boolean. Are stored procedure calls using the stored procedure escape syntax supported? True if so.
		$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_COMPARISONS, _                            ; 135 Returns a Boolean. Are subqueries in comparison expressions supported? True if so.
		$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_EXISTS, _                                 ; 136 Returns a Boolean. Are subqueries in "exists" expressions supported? True if so.
		$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_INS, _                                    ; 137 Returns a Boolean. Are subqueries in "in" statements supported? True if so.
		$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_QUANTIFIEDS, _                            ; 138 Returns a Boolean. Are subqueries in quantified expressions supported? True if so.
		$LOB_DBASE_META_SUPPORTS_TABLE_CORRELATION_NAMES, _                              ; 139 Returns a Boolean. Are table correlation names supported? True if so.
		$LOB_DBASE_META_SUPPORTS_TRANSACTION_ISOLATION_LEVEL, _                          ; 140 Returns a Boolean. Does this database support the given transaction isolation level? See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_SUPPORTS_TRANSACTIONS, _                                         ; 141 Returns a Boolean. Support the Database transactions? If not, invoking the method com::sun::star::sdbc::XConnection::commit() is a noop and the isolation level is TransactionIsolation_NONE. True if so.
		$LOB_DBASE_META_SUPPORTS_TYPE_CONVERSION, _                                      ; 142 Returns a Boolean. TRUE , if the Database supports the CONVERT function between SQL types, otherwise FALSE.
		$LOB_DBASE_META_SUPPORTS_UNION, _                                                ; 143 Returns a Boolean. Is SQL UNION supported? True if so.
		$LOB_DBASE_META_SUPPORTS_UNION_ALL, _                                            ; 144 Returns a Boolean. Is SQL UNION ALL supported? True if so.
		$LOB_DBASE_META_UPDATES_ARE_DETECTED, _                                          ; 145 Returns a Boolean. Indicates whether or not a visible row update can be detected by calling $LOB_RESULT_ROW_QUERY_IS_ROW_UPDATED. See _LOBase_DatabaseMetaDataQuery for Parameters. True if so.
		$LOB_DBASE_META_USES_LOCAL_FILE_PER_TABLE, _                                     ; 146 Returns a Boolean. Use the database one local file to save for each table. True if so.
		$LOB_DBASE_META_USES_LOCAL_FILES                                                 ; 147 Returns a Boolean. Use the database local files to save the tables. True if so.

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

; Text Horizontal Alignment
Global Const _                                                                           ; com.sun.star.style.ParagraphAdjust
		$LOB_PAR_TXT_ALIGN_HORI_LEFT = 0, _                                              ; Text is adjusted to the left border.
		$LOB_PAR_TXT_ALIGN_HORI_RIGHT = 1, _                                             ; Text is adjusted to the right border.
		$LOB_PAR_TXT_ALIGN_HORI_BLOCK = 2, _                                             ; Text is adjusted to both borders / stretched, except for last line.
		$LOB_PAR_TXT_ALIGN_HORI_CENTER = 3, _                                            ; Text is adjusted to the center.
		$LOB_PAR_TXT_ALIGN_HORI_STRETCH = 4                                              ; Text is adjusted to both borders / stretched, including last line.

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
		$LOB_REP_CON_TYPE_CHART = 1, _                                                   ; 1 A Chart
		$LOB_REP_CON_TYPE_FORMATTED_FIELD, _                                             ; 2 FORMATTED_FIELD
		$LOB_REP_CON_TYPE_IMAGE_CONTROL, _                                               ; 4 IMAGE_CONTROL
		$LOB_REP_CON_TYPE_LABEL, _                                                       ; 8 FIXED_TEXT
		$LOB_REP_CON_TYPE_LINE, _                                                        ; 16 A Vertical or Horizontal line.
		$LOB_REP_CON_TYPE_TEXT_BOX, _                                                    ; 32 A Text Box.
		$LOB_REP_CON_TYPE_ALL = 63                                                       ; 63 All of the above Control Types. (This value is the BitOR value of all above)

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
		$LOB_REP_OUTPUT_TYPE_UNKNOWN, _                                                  ; 0 The Output Document when the Report is executed is a unknown type.
		$LOB_REP_OUTPUT_TYPE_TEXT, _                                                     ; 1 The Output Document when the Report is executed will be a Text Document.
		$LOB_REP_OUTPUT_TYPE_SPREADSHEET                                                 ; 2 The Output Document when the Report is executed will be a Spreadsheet Document.

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
		$LOB_RESULT_CURSOR_MOVE_BEFORE_FIRST, _                                          ; 0 Moves before the first row.
		$LOB_RESULT_CURSOR_MOVE_FIRST, _                                                 ; 1 Moves to the first row.
		$LOB_RESULT_CURSOR_MOVE_PREVIOUS, _                                              ; 2 Moves back one row.
		$LOB_RESULT_CURSOR_MOVE_NEXT, _                                                  ; 3 Moves forward one row.
		$LOB_RESULT_CURSOR_MOVE_LAST, _                                                  ; 4 Moves to the last record.
		$LOB_RESULT_CURSOR_MOVE_AFTER_LAST, _                                            ; 5 Moves after the last record.
		$LOB_RESULT_CURSOR_MOVE_ABSOLUTE, _                                              ; 6 Moves to the row with the given row number.
		$LOB_RESULT_CURSOR_MOVE_RELATIVE                                                 ; 7 Moves backwards or forwards by the given amount: forwards for a positive value, and backwards for a negative value.

; Result Set Cursor Queries.
Global Enum _
		$LOB_RESULT_CURSOR_QUERY_IS_BEFORE_FIRST, _                                      ; 0 Is the cursor before the first record. This is the case if it has not yet been reset after entry.
		$LOB_RESULT_CURSOR_QUERY_IS_FIRST, _                                             ; 1 Is the cursor on the first entry.
		$LOB_RESULT_CURSOR_QUERY_IS_LAST, _                                              ; 2 Is the cursor on the last entry.
		$LOB_RESULT_CURSOR_QUERY_IS_AFTER_LAST, _                                        ; 3 Is the cursor after the last row when it is moved on with next.
		$LOB_RESULT_CURSOR_QUERY_GET_ROW                                                 ; 4 Retrieve the current row number.

; Column Nullability Constants.
Global Const _                                                                           ; com.sun.star.sdbc.ColumnValue Constant Group
		$LOB_RESULT_METADATA_COLUMN_NOT_NULLABLE = 0, _                                  ; The column does not allow NULL values.
		$LOB_RESULT_METADATA_COLUMN_NULLABLE = 1, _                                      ; The column does allow NULL values.
		$LOB_RESULT_METADATA_COLUMN_UNKNOWN_NULLABLE = 2                                 ; The nullability of the column is unknown.

; Column Metadata Query
Global Enum _
		$LOB_RESULT_METADATA_QUERY_GET_CATALOG_NAME, _                                   ; 0 Gets a column's table's catalog name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_DECIMAL_PLACE, _                                  ; 1 Gets a column's number of digits to right of the decimal point. Returns an Integer.
		$LOB_RESULT_METADATA_QUERY_GET_DISP_SIZE, _                                      ; 2 Gets the column's normal max width in chars. Returns an Integer.
		$LOB_RESULT_METADATA_QUERY_GET_LABEL, _                                          ; 3 Gets the suggested column title for use in printouts and displays. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_LENGTH, _                                         ; 4 Gets a column's number of decimal digits. Returns an Integer.
		$LOB_RESULT_METADATA_QUERY_GET_NAME, _                                           ; 5 Gets a column's name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_SCHEMA_NAME, _                                    ; 6 Gets a column's table's schema. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_TABLE_NAME, _                                     ; 7 Gets a column's table name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_GET_TYPE, _                                           ; 8 Gets the column's SQL type. Returns an Integer. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
		$LOB_RESULT_METADATA_QUERY_GET_TYPE_NAME, _                                      ; 9 Gets the column's database-specific type name. Returns a String.
		$LOB_RESULT_METADATA_QUERY_IS_AUTO_VALUE, _                                      ; 10 Query whether the column is automatically numbered, thus read-only (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_CASE_SENSITIVE, _                                  ; 11 Query whether a column's case matters (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_CURRENCY, _                                        ; 12 Query whether the column is a cash value (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_NULLABLE, _                                        ; 13 Query the nullability of values in the designated column. Returns an Integer. See Constants, $LOB_RESULT_METADATA_COLUMN_* as defined in LibreOfficeBase_Constants.au3.
		$LOB_RESULT_METADATA_QUERY_IS_READ_ONLY, _                                       ; 14 Query whether a column is definitely not writable (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_SEARCHABLE, _                                      ; 15 Query whether the column can be used in a where clause (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_SIGNED, _                                          ; 16 Query whether values in the column are signed numbers (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_WRITABLE, _                                        ; 17 Query whether it is possible for a write on the column to succeed (True if so). Returns a Boolean.
		$LOB_RESULT_METADATA_QUERY_IS_WRITABLE_DEFINITE                                  ; 18 Query whether a write on the column will definitely succeed (True if so). Returns a Boolean.

; Result Set Row Modification Commands.
Global Enum _
		$LOB_RESULT_ROW_MOD_NULL, _                                                      ; 0 Sets the column content to NULL
		$LOB_RESULT_ROW_MOD_BOOL, _                                                      ; 1 Changes the content of specified column to the called logical value.
		$LOB_RESULT_ROW_MOD_BYTE, _                                                      ; 2 Changes the content of specified column to the called byte.
		$LOB_RESULT_ROW_MOD_SHORT, _                                                     ; 3 Changes the content of specified column to the called integer.
		$LOB_RESULT_ROW_MOD_INT, _                                                       ; 4 Changes the content of specified column to the called integer.
		$LOB_RESULT_ROW_MOD_LONG, _                                                      ; 5 Changes the content of specified column to the called integer.
		$LOB_RESULT_ROW_MOD_FLOAT, _                                                     ; 6 Changes the content of specified column to the called decimal number.
		$LOB_RESULT_ROW_MOD_DOUBLE, _                                                    ; 7 Changes the content of specified column to the called decimal number.
		$LOB_RESULT_ROW_MOD_STRING, _                                                    ; 8 Changes the content of specified column to the called string.
		$LOB_RESULT_ROW_MOD_BYTES, _                                                     ; 9 Changes the content of specified column to the called byte array.
		$LOB_RESULT_ROW_MOD_DATE, _                                                      ; 10 Changes the content of specified column to the called date.
		$LOB_RESULT_ROW_MOD_TIME, _                                                      ; 11 Changes the content of specified column to the called time.
		$LOB_RESULT_ROW_MOD_TIMESTAMP                                                    ; 12 Changes the content of specified column to the called Date and Time (Timestamp).

; Result Set Queries.
Global Enum _
		$LOB_RESULT_ROW_QUERY_IS_ROW_INSERTED, _                                         ; 0 Indicates if this is a new row.
		$LOB_RESULT_ROW_QUERY_IS_ROW_UPDATED, _                                          ; 1 Indicates if the current row has been altered.
		$LOB_RESULT_ROW_QUERY_IS_ROW_DELETED                                             ; 2 Indicates if the current row has been deleted.

; Result Set Row Read Commands.
Global Enum _
		$LOB_RESULT_ROW_READ_STRING, _                                                   ; 0 Returns the content of the column as a character string.
		$LOB_RESULT_ROW_READ_BOOL, _                                                     ; 1 Returns the content of the column as a boolean value.
		$LOB_RESULT_ROW_READ_BYTE, _                                                     ; 2 Returns the content of the column as a single byte.
		$LOB_RESULT_ROW_READ_SHORT, _                                                    ; 3 Returns the content of the column as an integer.
		$LOB_RESULT_ROW_READ_INT, _                                                      ; 4 Returns the content of the column as an integer.
		$LOB_RESULT_ROW_READ_LONG, _                                                     ; 5 Returns the content of the column as an integer.
		$LOB_RESULT_ROW_READ_FLOAT, _                                                    ; 6 Returns the content of the column as a single precision decimal number.
		$LOB_RESULT_ROW_READ_DOUBLE, _                                                   ; 7 Returns the content of the column as a double precision decimal number.
		$LOB_RESULT_ROW_READ_BYTES, _                                                    ; 8 Returns the content of the column as an array of single bytes.
		$LOB_RESULT_ROW_READ_DATE, _                                                     ; 9 Returns the content of the column as a date.
		$LOB_RESULT_ROW_READ_TIME, _                                                     ; 10 Returns the content of the column as a time value.
		$LOB_RESULT_ROW_READ_TIMESTAMP, _                                                ; 11 Returns the content of the column as a timestamp (date and time).
		$LOB_RESULT_ROW_READ_WAS_NULL                                                    ; 12 Indicates if the value of the most recently read column was NULL.

; Result Set Row Update Commands.
Global Enum _
		$LOB_RESULT_ROW_UPDATE_INSERT, _                                                 ; 0 Saves a new row.
		$LOB_RESULT_ROW_UPDATE_UPDATE, _                                                 ; 1 Confirms alteration of the current row.
		$LOB_RESULT_ROW_UPDATE_DELETE, _                                                 ; 2 Deletes the current row.
		$LOB_RESULT_ROW_UPDATE_CANCEL_UPDATE, _                                          ; 3 Reverses changes in the current row.
		$LOB_RESULT_ROW_UPDATE_MOVE_TO_INSERT, _                                         ; 4 Moves the cursor into a row corresponding to a new record.
		$LOB_RESULT_ROW_UPDATE_MOVE_TO_CURRENT                                           ; 5 After the entry of a new record, returns the cursor to its previous position.

; Database Result Set Type Constants
Global Const _                                                                           ; com.sun.star.sdbc.ResultSetType Constant Group
		$LOB_RESULT_TYPE_FORWARD_ONLY = 1003, _                                          ; The Result Set cursor may move only forward.
		$LOB_RESULT_TYPE_SCROLL_INSENSITIVE = 1004, _                                    ; The Result Set is scrollable but generally not sensitive to changes made by others.
		$LOB_RESULT_TYPE_SCROLL_SENSITIVE = 1005                                         ; The Result Set is scrollable and generally sensitive to changes made by others.

; Database SubComponent type.
Global Const _                                                                           ; com.sun.star.sdb.application.DatabaseObject
		$LOB_SUB_COMP_TYPE_TABLE = 0, _                                                  ; A Database Table.
		$LOB_SUB_COMP_TYPE_QUERY = 1, _                                                  ; A Database Query.
		$LOB_SUB_COMP_TYPE_FORM = 2, _                                                   ; A Database Form.
		$LOB_SUB_COMP_TYPE_REPORT = 3                                                    ; A Database Report.
