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
; Description ...: Provides basic functionality through AutoIt for Registering, Unregistering, and connecting to, etc. L.O. Base Databases.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_DatabaseAutoCommit
; _LOBase_DatabaseCommit
; _LOBase_DatabaseConnectionClose
; _LOBase_DatabaseConnectionGet
; _LOBase_DatabaseGetDefaultQuote
; _LOBase_DatabaseGetObjByDoc
; _LOBase_DatabaseGetObjByURL
; _LOBase_DatabaseIsReadOnly
; _LOBase_DatabaseMetaDataQuery
; _LOBase_DatabaseName
; _LOBase_DatabaseRegisteredAdd
; _LOBase_DatabaseRegisteredExists
; _LOBase_DatabaseRegisteredGetNames
; _LOBase_DatabaseRegisteredRemoveByName
; _LOBase_DatabaseRequiresPassword
; _LOBase_DatabaseRollback
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseAutoCommit
; Description ...: Set or Retrieve the Database's current AutoCommit setting. See remarks.
; Syntax ........: _LOBase_DatabaseAutoCommit(ByRef $oConnection[, $bAutoCommit = Null])
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $bAutoCommit         - [optional] a boolean value. Default is Null. If True, all of the SQL statements will be executed and committed as individual transactions.
; Return values .: Success: 1 or Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a connection Object.
;                  @Error 1 @Extended 3 Return 0 = $bAutoCommit not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bAutoCommit
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Boolean = Success. All optional parameters were called with Null, returning current AutoCommit setting as a Boolean value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: By default, new connections have auto-commit active.
;                  You can only modify the AutoCommit setting on private connections, however you can retrieve the current setting of AutoCommit for non-private or private connections.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseAutoCommit(ByRef $oConnection, $bAutoCommit = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LO_VarsAreNull($bAutoCommit) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oConnection.getAutoCommit())

	If Not IsBool($bAutoCommit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oConnection.setAutoCommit($bAutoCommit)
	If Not ($oConnection.getAutoCommit() = $bAutoCommit) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DatabaseAutoCommit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseCommit
; Description ...: Commit any changes made to the database since the last save.
; Syntax ........: _LOBase_DatabaseCommit(ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a connection Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully executed commit command.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is unnecessary if AutoCommit is active (default).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseCommit(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oConnection.commit()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DatabaseCommit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseConnectionClose
; Description ...: Close an opened Database connection.
; Syntax ........: _LOBase_DatabaseConnectionClose(ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a connection Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Connection is already closed.
;                  @Error 3 @Extended 2 Return 0 = Failed to close connection.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully closed the connection.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseConnectionGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseConnectionClose(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oConnection.Close()

	If Not $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DatabaseConnectionClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseConnectionGet
; Description ...: Create a connection to a Database.
; Syntax ........: _LOBase_DatabaseConnectionGet(ByRef $oDBase[, $sUser = ""[, $sPass = ""[, $bPromptUser = False[, $bPrivate = False]]]])
; Parameters ....: $oDBase              - [in/out] an object. A Database object returned by a previous _LOBase_DatabaseGetObjByDoc or _LOBase_DatabaseGetObjByURL function.
;                  $sUser               - [optional] a string value. Default is "". The Username for connecting to the Database. If none, leave as a blank string.
;                  $sPass               - [optional] a string value. Default is "". The Password for connecting to the Database. If none, leave as a blank string.
;                  $bPromptUser         - [optional] a boolean value. Default is False. If True, $sUser and $sPass are ignored, and the user is prompted for the required information.
;                  $bPrivate            - [optional] a boolean value. Default is False. If True, a private connection is created, otherwise a public connection is created.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDBase not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sUser not a String.
;                  @Error 1 @Extended 3 Return 0 = $sPass not a String.
;                  @Error 1 @Extended 4 Return 0 = $bPromptUser not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bPrivate not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.task.InteractionHandler" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to Connect to the Database.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully connected to the Database, returning a Connection Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseConnectionClose
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseConnectionGet(ByRef $oDBase, $sUser = "", $sPass = "", $bPromptUser = False, $bPrivate = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBConnection, $oHandler

	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sUser) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sPass) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bPromptUser) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bPrivate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If $bPrivate Then
		If $bPromptUser Then
			$oServiceManager = __LO_ServiceManager()
			If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oHandler = $oServiceManager.createInstance("com.sun.star.task.InteractionHandler")
			If Not IsObj($oHandler) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$oDBConnection = $oDBase.getIsolatedConnectionWithCompletion($oHandler)

		Else
			$oDBConnection = $oDBase.getIsolatedConnection($sUser, $sPass)
		EndIf

	Else
		If $bPromptUser Then
			$oServiceManager = __LO_ServiceManager()
			If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			$oHandler = $oServiceManager.createInstance("com.sun.star.task.InteractionHandler")
			If Not IsObj($oHandler) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$oDBConnection = $oDBase.ConnectWithCompletion($oHandler)

		Else
			$oDBConnection = $oDBase.getConnection($sUser, $sPass)
		EndIf
	EndIf

	If Not IsObj($oDBConnection) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDBConnection)
EndFunc   ;==>_LOBase_DatabaseConnectionGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseGetDefaultQuote
; Description ...: Retrieves the string used to quote SQL identifiers.
; Syntax ........: _LOBase_DatabaseGetDefaultQuote(ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a connection Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Quotation character.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the default Quote character used by the Database to quote SQL identifiers.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This returns a space " " if identifier quoting is not supported.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseGetDefaultQuote(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sQuote

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sQuote = $oConnection.MetaData.getIdentifierQuoteString()
	If Not IsString($sQuote) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sQuote)
EndFunc   ;==>_LOBase_DatabaseGetDefaultQuote

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseGetObjByDoc
; Description ...: Retrieve a Database Object from a Document Object.
; Syntax ........: _LOBase_DatabaseGetObjByDoc(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Document Save path.
;                  @Error 3 @Extended 2 Return 0 = Document hasn't been saved yet.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Database Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Database Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Document must be saved first, before you can obtain a Database Object for it.
;                  If you intend to move or delete the Database Document file after retrieving a Database Object, you must allow the function that created the Database Object to return, and the Database Object to be released by AutoIt (only works for Local variable, by the way). Otherwise the file will be shown as "In Use" and cannot be moved or deleted until AutoIt releases the Object. See AutoIt Help file "Obj/COM Reference", Last paragraph in section "An example usage of COM in AutoIt"
; Related .......: _LOBase_DatabaseGetObjByURL
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseGetObjByDoc(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext, $oDBase
	Local $sURL

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sURL = $oDoc.URL()
	If Not IsString($sURL) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($sURL = "") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oDBase = $oDBaseContext.getByName($sURL)
	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDBase)
EndFunc   ;==>_LOBase_DatabaseGetObjByDoc

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseGetObjByURL
; Description ...: Retrieve a Database Object from a URL or registered Database name.
; Syntax ........: _LOBase_DatabaseGetObjByURL($sURL)
; Parameters ....: $sURL                - a string value. The File path of the Database file or a Database name that is registered.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sURL not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to convert called path to URL format.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Database Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning requested Database Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving a Database Object this way allows you to edit a Database without using a GUI.
;                  If you intend to move or delete the Database Document file after retrieving a Database Object, you must allow the function that created the Database Object to return, and the Database Object to be released by AutoIt (only works for Local variable, by the way). Otherwise the file will be shown as "In Use" and cannot be moved or deleted until AutoIt releases the Object. See AutoIt Help file "Obj/COM Reference", Last paragraph in section "An example usage of COM in AutoIt"
; Related .......: _LOBase_DatabaseGetObjByDoc
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseGetObjByURL($sURL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext, $oDBase
	Local $sFileURL = $sURL

	If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If StringInStr($sFileURL, "\") Then $sFileURL = _LO_PathConvert($sFileURL, $LO_PATHCONV_OFFICE_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oDBase = $oDBaseContext.getByName($sFileURL)
	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDBase)
EndFunc   ;==>_LOBase_DatabaseGetObjByURL

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseIsReadOnly
; Description ...: Check if a Database is currently in a Read-Only state.
; Syntax ........: _LOBase_DatabaseIsReadOnly(ByRef $oDBase)
; Parameters ....: $oDBase              - [in/out] an object. A Database object returned by a previous _LOBase_DatabaseGetObjByDoc or _LOBase_DatabaseGetObjByURL function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDBase not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query Database for Read-Only Status.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If Database is currently Read-Only, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseRequiresPassword
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseIsReadOnly(ByRef $oDBase)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn

	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bReturn = $oDBase.DatabaseDocument.DataSource.IsReadOnly()
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_DatabaseIsReadOnly

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseMetaDataQuery
; Description ...: Query a Database's MetaData.
; Syntax ........: _LOBase_DatabaseMetaDataQuery(ByRef $oConnection, $iQuery[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null[, $vParam6 = Null]]]]]])
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
;                  $iQuery              - an integer value (0-148). The query to perform. See Constants, $LOB_DBASE_META_* as defined in LibreOfficeBase_Constants.au3.
;                  $vParam1             - [optional] a variant value. Default is Null. The first Parameter required by the Query. See remarks for the queries that have parameters.
;                  $vParam2             - [optional] a variant value. Default is Null. The second Parameter required by the Query. See remarks for the queries that have parameters.
;                  $vParam3             - [optional] a variant value. Default is Null. The third Parameter required by the Query. See remarks for the queries that have parameters.
;                  $vParam4             - [optional] a variant value. Default is Null. The fourth Parameter required by the Query. See remarks for the queries that have parameters.
;                  $vParam5             - [optional] a variant value. Default is Null. The fifth Parameter required by the Query. See remarks for the queries that have parameters.
;                  $vParam6             - [optional] a variant value. Default is Null. The sixth Parameter required by the Query. See remarks for the queries that have parameters.
; Return values .: Success: Variable
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a connection Object.
;                  @Error 1 @Extended 3 Return 0 = $iQuery not an Integer, less than 0 or greater than last Query command. See Constants, $LOB_DBASE_META_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $vParam1 not an Integer, less than 1003 or 1005. See Constants, $LOB_RESULT_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $vParam1 not an Integer, less than -16 or 2014. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $vParam2 not an Integer, less than -16 or 2014. See Constants, $LOB_DATA_TYPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $vParam2 not an Integer, less than 1007 or 1008. See Constants, $LOB_DBASE_RESULT_SET_CONCURRENCY_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $vParam1 not an Integer, less than 0 or 8. See Constants, $LOB_DBASE_TRANSACTION_ISOLATION_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $vParam1 not a String.
;                  @Error 1 @Extended 10 Return 0 = $vParam2 not a String.
;                  @Error 1 @Extended 11 Return 0 = $vParam3 not a String.
;                  @Error 1 @Extended 12 Return 0 = $vParam4 not a String.
;                  @Error 1 @Extended 13 Return 0 = $vParam4 not an Array.
;                  @Error 1 @Extended 14 Return 0 = $vParam4 not an Integer, less than 2000 or 2002. See Constants, LOB_DATA_TYPE_OBJECT, $LOB_DATA_TYPE_DISTINCT, $LOB_DATA_TYPE_STRUCT, as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 15 Return 0 = $vParam4 not an Integer, less than 0 or 2. See Constants, $LOB_DBASE_BEST_ROW_SCOPE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 16 Return 0 = $vParam4 not a Boolean.
;                  @Error 1 @Extended 17 Return 0 = $vParam5 not a Boolean.
;                  @Error 1 @Extended 18 Return 0 = $vParam5 not a String.
;                  @Error 1 @Extended 19 Return 0 = $vParam6 not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to obtain Query command.
;                  @Error 3 @Extended 2 Return 0 = Connection is already closed.
;                  @Error 3 @Extended 3 Return 0 = Failed to perform Query.
;                  --Success--
;                  @Error 0 @Extended 0 Return Variable = Success. Returning the Result of the query. See respective query for expected return.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Parameters that require a String, such as a Table name etc., accept certain wildcards also: Within a pattern String, "%" means match any substring of 0 or more characters, and "_" means match any one character.
;                  Some queries requires parameters to call them. See below for a list of queries with parameters, how many, and which $vParam to call which data type in.
;                  $LOB_DBASE_META_DELETES_ARE_DETECTED Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer.
;                  $LOB_DBASE_META_GET_BEST_ROW_ID Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA name, "" retrieves those without a schema. String; $vParam3 = TABLE name. String; $vParam4 = SCOPE, the scope of interest, Use Constants $LOB_DBASE_BEST_ROW_SCOPE_*. Integer; $vParam5 = NULLABLE, include columns that are Nullable? Boolean value.
;                  $LOB_DBASE_META_GET_COL_PRIVILEGES Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA name, "" retrieves those without a schema. String; $vParam3 = TABLE name. String; $vParam4 = COLUMN NAME PATTERN. String.
;                  $LOB_DBASE_META_GET_COLS Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA name, "" retrieves those without a schema. String; $vParam3 = TABLE name. String; $vParam4 = COLUMN NAME PATTERN. String.
;                  $LOB_DBASE_META_GET_CROSS_REF Call $vParam1 with PRIMARY CATALOG name, "" retrieves those without a catalog. String; $vParam2 = PRIMARY SCHEMA name, "" retrieves those without a schema. String; $vParam3 = PRIMARY TABLE name. String; $vParam4 = FOREIGN CATALOG name, "" retrieves those without a catalog. String; $vParam5 = FOREIGN SCHEMA name, "" retrieves those without a schema. String; $vParam6 = FOREIGN TABLE name. String;
;                  $LOB_DBASE_META_GET_EXPORTED_KEYS Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA name, "" retrieves those without a schema. String; $vParam3 = TABLE name. String.
;                  $LOB_DBASE_META_GET_IMPORTED_KEYS Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA name, "" retrieves those without a schema. String; $vParam3 = TABLE name. String.
;                  $LOB_DBASE_META_GET_INDEX_INFO Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA name, "" retrieves those without a schema. String; $vParam3 = TABLE name. String; $vParam4 = UNIQUE, when TRUE, return only indices for unique values; when FALSE, return indices regardless of whether unique or not. Boolean value.; $vParam5 = APPROXIMATE, when TRUE, result is allowed to reflect approximate or out of data values; when FALSE, results are requested to be accurate. Boolean value.
;                  $LOB_DBASE_META_GET_PRIMARY_KEY Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA name, "" retrieves those without a schema. String; $vParam3 = TABLE name. String.
;                  $LOB_DBASE_META_GET_PROCEDURE_COLS Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA PATTERN, a schema name pattern; "" retrieves those without a schema. String; $vParam3 = PROCEDURE NAME PATTERN. String; $vParam4 = COLUMN NAME PATTERN. String.
;                  $LOB_DBASE_META_GET_PROCEDURES Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA PATTERN, a schema name pattern; "" retrieves those without a schema. String; $vParam3 = PROCEDURE NAME PATTERN. String.
;                  $LOB_DBASE_META_GET_TABLE_PRIVILEGES Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA PATTERN, a schema name pattern; "" retrieves those without a schema. String; $vParam3 = TABLE NAME PATTERN. String.
;                  $LOB_DBASE_META_GET_TABLES Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA PATTERN, a schema name pattern; "" retrieves those without a schema. String; $vParam3 = TABLE NAME PATTERN. String; $vParam4 = TYPES, Array containing Table Types to Include. Array. (See $LOB_DBASE_META_GET_TABLE_TYPES)
;                  $LOB_DBASE_META_GET_UDTS Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA PATTERN, a schema name pattern; "" retrieves those without a schema. String; $vParam3 = TYPE NAME PATTERN, a type name pattern; may be a fully-qualified name. String; $vParam4 = TYPES, Accepts Constants, $LOB_DATA_TYPE_ Object, Struct and Distinct. Integer.
;                  $LOB_DBASE_META_GET_VERSION_COLS Call $vParam1 with CATALOG name, "" retrieves those without a catalog. String; $vParam2 = SCHEMA name, "" retrieves those without a schema. String; $vParam3 = TABLE name. String.
;                  $LOB_DBASE_META_INSERTS_ARE_DETECTED Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer.
;                  $LOB_DBASE_META_OTHERS_DELETES_ARE_VISIBLE Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer.
;                  $LOB_DBASE_META_OTHERS_INSERTS_ARE_VISIBLE Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer.
;                  $LOB_DBASE_META_OTHERS_UPDATES_ARE_VISIBLE Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer.
;                  $LOB_DBASE_META_OWN_DELETES_ARE_VISIBLE Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer.
;                  $LOB_DBASE_META_OWN_INSERTS_ARE_VISIBLE Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer.
;                  $LOB_DBASE_META_OWN_UPDATES_ARE_VISIBLE Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer.
;                  $LOB_DBASE_META_SUPPORTS_CONVERT Call $vParam1 with one of the Data Type constants, $LOB_DATA_TYPE_*. This is the FROM type. Integer. $vParam2 = TO type, one of the Data Type constants, $LOB_DATA_TYPE_*. Integer.
;                  $LOB_DBASE_META_SUPPORTS_RESULT_SET_CONCURRENCY Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer; $vParam2 = CONCURRENCY, one of the Result Set Concurrency constants, $LOB_DBASE_RESULT_SET_CONCURRENCY_*. Integer.
;                  $LOB_DBASE_META_SUPPORTS_RESULT_SET_TYPE Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer.
;                  $LOB_DBASE_META_SUPPORTS_TRANSACTION_ISOLATION_LEVEL Call $vParam1 with one of the Transaction Isolation Level constants, $LOB_DBASE_TRANSACTION_ISOLATION_*. Integer.
;                  $LOB_DBASE_META_UPDATES_ARE_DETECTED Call $vParam1 with one of the Result Set Type constants, $LOB_RESULT_TYPE_*. Integer.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseMetaDataQuery(ByRef $oConnection, $iQuery, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null, $vParam6 = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $sCall

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iQuery, $LOB_DBASE_META_ALL_PROCEDURES_ARE_CALLABLE, $LOB_DBASE_META_USES_LOCAL_FILES) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$sCall = __LOBase_DatabaseMetaGetQuery($iQuery)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oConnection.isClosed() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Switch $iQuery
		Case $LOB_DBASE_META_ALL_PROCEDURES_ARE_CALLABLE, $LOB_DBASE_META_ALL_TABLES_ARE_SELECTABLE, $LOB_DBASE_META_DATA_DEFINITION_CAUSES_TRANSACTION_COMMIT, _
				$LOB_DBASE_META_DATA_DEFINITION_IGNORED_IN_TRANSACTIONS, $LOB_DBASE_META_DOES_MAX_ROW_SIZE_INCLUDE_BLOBS, $LOB_DBASE_META_IS_CATALOG_AT_START, _
				$LOB_DBASE_META_IS_READ_ONLY, $LOB_DBASE_META_NULL_PLUS_NON_NULL_IS_NULL, $LOB_DBASE_META_NULLS_ARE_SORTED_AT_END, $LOB_DBASE_META_NULLS_ARE_SORTED_AT_START, _
				$LOB_DBASE_META_NULLS_ARE_SORTED_HIGH, $LOB_DBASE_META_NULLS_ARE_SORTED_LOW, $LOB_DBASE_META_STORES_LOWER_CASE_IDS, $LOB_DBASE_META_STORES_MIXED_CASE_IDS, _
				$LOB_DBASE_META_STORES_UPPER_CASE_IDS, $LOB_DBASE_META_STORES_LOWER_CASE_QUOTED_IDS, $LOB_DBASE_META_STORES_MIXED_CASE_QUOTED_IDS, _
				$LOB_DBASE_META_STORES_UPPER_CASE_QUOTED_IDS, $LOB_DBASE_META_SUPPORTS_ALTER_TABLE_WITH_ADD_COL, $LOB_DBASE_META_SUPPORTS_ALTER_TABLE_WITH_DROP_COL, _
				$LOB_DBASE_META_SUPPORTS_ANSI92_ENTRY_LEVEL_SQL, $LOB_DBASE_META_SUPPORTS_ANSI92_FULL_SQL, $LOB_DBASE_META_SUPPORTS_ANSI92_INTERMEDIATE_SQL, _
				$LOB_DBASE_META_SUPPORTS_BATCH_UPDATES, $LOB_DBASE_META_SUPPORTS_CATALOGS_IN_DATA_MANIPULATION, $LOB_DBASE_META_SUPPORTS_CATALOGS_IN_INDEX_DEFINITIONS, _
				$LOB_DBASE_META_SUPPORTS_CATALOGS_IN_PRIVILEGE_DEFINITIONS, $LOB_DBASE_META_SUPPORTS_CATALOGS_IN_PROCEDURE_CALLS, $LOB_DBASE_META_SUPPORTS_CATALOGS_IN_TABLE_DEFINITION, _
				$LOB_DBASE_META_SUPPORTS_COL_ALIASING, $LOB_DBASE_META_SUPPORTS_CORE_SQL_GRAMMAR, $LOB_DBASE_META_SUPPORTS_CORRELATED_SUBQUERIES, _
				$LOB_DBASE_META_SUPPORTS_DATA_DEFINITION_AND_DATA_MANIPULATION_TRANSACTIONS, $LOB_DBASE_META_SUPPORTS_DATA_MANIPULATION_TRANSACTIONS_ONLY, _
				$LOB_DBASE_META_SUPPORTS_DIFF_TABLE_CORRELATION_NAMES, $LOB_DBASE_META_SUPPORTS_EXPRESSIONS_IN_ORDER_BY, $LOB_DBASE_META_SUPPORTS_EXTENDED_SQL_GRAMMAR, _
				$LOB_DBASE_META_SUPPORTS_FULL_OUTER_JOINS, $LOB_DBASE_META_SUPPORTS_GROUP_BY, $LOB_DBASE_META_SUPPORTS_GROUP_BY_BEYOND_SELECT, _
				$LOB_DBASE_META_SUPPORTS_GROUP_BY_UNRELATED, $LOB_DBASE_META_SUPPORTS_INTEGRITY_ENHANCMENT_FACILITY, $LOB_DBASE_META_SUPPORTS_LIKE_ESCAPE_CLAUSE, _
				$LOB_DBASE_META_SUPPORTS_LIMITED_OUTER_JOINS, $LOB_DBASE_META_SUPPORTS_MINIMUM_SQL_GRAMMAR, $LOB_DBASE_META_SUPPORTS_MIXED_CASE_IDS, _
				$LOB_DBASE_META_SUPPORTS_MIXED_CASE_QUOTED_IDS, $LOB_DBASE_META_SUPPORTS_MULTIPLE_RESULT_SETS, $LOB_DBASE_META_SUPPORTS_MULTIPLE_TRANSACTIONS, _
				$LOB_DBASE_META_SUPPORTS_NON_NULLABLE_COLS, $LOB_DBASE_META_SUPPORTS_OPEN_CURSORS_ACROSS_COMMIT, $LOB_DBASE_META_SUPPORTS_OPEN_CURSORS_ACROSS_ROLLBACK, _
				$LOB_DBASE_META_SUPPORTS_OPEN_STATEMENTS_ACROSS_COMMIT, $LOB_DBASE_META_SUPPORTS_OPEN_STATEMENTS_ACROSS_ROLLBACK, _
				$LOB_DBASE_META_SUPPORTS_ORDER_BY_UNRELATED, $LOB_DBASE_META_SUPPORTS_OUTER_JOINS, $LOB_DBASE_META_SUPPORTS_POSITIONED_DELETE, _
				$LOB_DBASE_META_SUPPORTS_POSITIONED_UPDATE, $LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_DATA_MANIPULATION, $LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_INDEX_DEFINITIONS, _
				$LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_PRIVILEGE_DEFINITIONS, $LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_PROCEDURE_CALLS, $LOB_DBASE_META_SUPPORTS_SCHEMAS_IN_TABLE_DEFINITION, _
				$LOB_DBASE_META_SUPPORTS_SELECT_FOR_UPDATE, $LOB_DBASE_META_SUPPORTS_STORED_PROCEDURES, $LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_COMPARISONS, _
				$LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_EXISTS, $LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_INS, $LOB_DBASE_META_SUPPORTS_SUBQUERIES_IN_QUANTIFIEDS, _
				$LOB_DBASE_META_SUPPORTS_TABLE_CORRELATION_NAMES, $LOB_DBASE_META_SUPPORTS_TRANSACTIONS, $LOB_DBASE_META_SUPPORTS_TYPE_CONVERSION, _
				$LOB_DBASE_META_SUPPORTS_UNION, $LOB_DBASE_META_SUPPORTS_UNION_ALL, $LOB_DBASE_META_USES_LOCAL_FILE_PER_TABLE, $LOB_DBASE_META_USES_LOCAL_FILES

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "()")
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_GET_CATALOG_SEPARATOR, $LOB_DBASE_META_GET_CATALOG_TERM, $LOB_DBASE_META_GET_DATABASE_PRODUCT_NAME, $LOB_DBASE_META_GET_DATABASE_PRODUCT_VERSION, _
				$LOB_DBASE_META_GET_DRIVER_NAME, $LOB_DBASE_META_GET_EXTRA_NAME_CHARS, $LOB_DBASE_META_GET_IDENTIFIER_QUOTE_STRING, $LOB_DBASE_META_GET_NUMERIC_FUNCS, _
				$LOB_DBASE_META_GET_PROCEDURE_TERM, $LOB_DBASE_META_GET_SCHEMA_TERM, $LOB_DBASE_META_GET_SEARCH_STRING_ESCAPE, $LOB_DBASE_META_GET_SQL_KEYWORDS, _
				$LOB_DBASE_META_GET_STRING_FUNCS, $LOB_DBASE_META_GET_SYSTEM_FUNCS, $LOB_DBASE_META_GET_TIME_DATE_FUNCS, $LOB_DBASE_META_GET_URL, _
				$LOB_DBASE_META_GET_USERNAME

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "()")
			If Not IsString($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_GET_DEFAULT_TRANSACTION_ISOLATION, $LOB_DBASE_META_GET_DRIVER_MAJOR_VERSION, $LOB_DBASE_META_GET_DRIVER_MINOR_VERSION, _
				$LOB_DBASE_META_GET_DRIVER_VERSION, $LOB_DBASE_META_GET_MAX_BINARY_LITERAL_LEN, $LOB_DBASE_META_GET_MAX_CATALOG_NAME_LEN, _
				$LOB_DBASE_META_GET_MAX_CHAR_LITERAL_LEN, $LOB_DBASE_META_GET_MAX_COL_NAME_LEN, $LOB_DBASE_META_GET_MAX_COLS_IN_GROUP_BY, _
				$LOB_DBASE_META_GET_MAX_COLS_IN_INDEX, $LOB_DBASE_META_GET_MAX_COLS_IN_ORDER_BY, $LOB_DBASE_META_GET_MAX_COLS_IN_SEL, _
				$LOB_DBASE_META_GET_MAX_COLS_IN_TABLE, $LOB_DBASE_META_GET_MAX_CONNECTIONS, $LOB_DBASE_META_GET_MAX_CURSOR_NAME_LEN, _
				$LOB_DBASE_META_GET_MAX_INDEX_LEN, $LOB_DBASE_META_GET_MAX_PROCEDURE_NAME_LEN, $LOB_DBASE_META_GET_MAX_ROW_SIZE, _
				$LOB_DBASE_META_GET_MAX_SCHEMA_NAME_LEN, $LOB_DBASE_META_GET_MAX_STATEMENT_LEN, $LOB_DBASE_META_GET_MAX_STATEMENTS, _
				$LOB_DBASE_META_GET_MAX_TABLE_NAME_LEN, $LOB_DBASE_META_GET_MAX_TABLES_IN_SEL, $LOB_DBASE_META_GET_MAX_USER_NAME_LEN

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "()")
			If Not IsInt($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_GET_CATALOGS, $LOB_DBASE_META_GET_SCHEMAS, $LOB_DBASE_META_GET_TABLE_TYPES, $LOB_DBASE_META_GET_TYPE_INFO
			$vReturn = Execute("$oConnection.MetaData" & $sCall & "()")
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_DELETES_ARE_DETECTED, $LOB_DBASE_META_INSERTS_ARE_DETECTED, $LOB_DBASE_META_OTHERS_DELETES_ARE_VISIBLE, _
				$LOB_DBASE_META_OTHERS_INSERTS_ARE_VISIBLE, $LOB_DBASE_META_OTHERS_UPDATES_ARE_VISIBLE, $LOB_DBASE_META_OWN_DELETES_ARE_VISIBLE, _
				$LOB_DBASE_META_OWN_INSERTS_ARE_VISIBLE, $LOB_DBASE_META_OWN_UPDATES_ARE_VISIBLE, $LOB_DBASE_META_SUPPORTS_RESULT_SET_TYPE, _
				$LOB_DBASE_META_UPDATES_ARE_DETECTED

			If Not __LO_IntIsBetween($vParam1, $LOB_RESULT_TYPE_FORWARD_ONLY, $LOB_RESULT_TYPE_SCROLL_SENSITIVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ")")
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_SUPPORTS_CONVERT
			If Not __LO_IntIsBetween($vParam1, $LOB_DATA_TYPE_LONGNVARCHAR, $LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
			If Not __LO_IntIsBetween($vParam2, $LOB_DATA_TYPE_LONGNVARCHAR, $LOB_DATA_TYPE_TIMESTAMP_WITH_TIMEZONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ", " & $vParam2 & ")")
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_SUPPORTS_RESULT_SET_CONCURRENCY
			If Not __LO_IntIsBetween($vParam1, $LOB_RESULT_TYPE_FORWARD_ONLY, $LOB_RESULT_TYPE_SCROLL_SENSITIVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
			If Not __LO_IntIsBetween($vParam2, $LOB_DBASE_RESULT_SET_CONCURRENCY_READ_ONLY, $LOB_DBASE_RESULT_SET_CONCURRENCY_UPDATABLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ", " & $vParam2 & ")")
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_SUPPORTS_TRANSACTION_ISOLATION_LEVEL
			If Not __LO_IntIsBetween($vParam1, $LOB_DBASE_TRANSACTION_ISOLATION_NONE, $LOB_DBASE_TRANSACTION_ISOLATION_SERIALIZED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ")")
			If Not IsBool($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_GET_EXPORTED_KEYS, $LOB_DBASE_META_GET_IMPORTED_KEYS, $LOB_DBASE_META_GET_PRIMARY_KEY, $LOB_DBASE_META_GET_PROCEDURES, _
				$LOB_DBASE_META_GET_TABLE_PRIVILEGES, $LOB_DBASE_META_GET_VERSION_COLS

			If Not IsString($vParam1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
			If Not IsString($vParam2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
			If Not IsString($vParam3) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ", " & $vParam2 & ", " & $vParam3 & ", " & ")")
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_GET_COLS, $LOB_DBASE_META_GET_COL_PRIVILEGES, $LOB_DBASE_META_GET_PROCEDURE_COLS
			If Not IsString($vParam1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
			If Not IsString($vParam2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
			If Not IsString($vParam3) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
			If Not IsString($vParam4) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ", " & $vParam2 & ", " & $vParam3 & ", " & $vParam4 & ")")
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_GET_TABLES
			If Not IsString($vParam1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
			If Not IsString($vParam2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
			If Not IsString($vParam3) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
			If Not IsArray($vParam4) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ", " & $vParam2 & ", " & $vParam3 & ", " & $vParam4 & ")")
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_GET_UDTS
			If Not IsString($vParam1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
			If Not IsString($vParam2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
			If Not IsString($vParam3) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
			If Not __LO_IntIsBetween($vParam4, $LOB_DATA_TYPE_OBJECT, $LOB_DATA_TYPE_STRUCT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ", " & $vParam2 & ", " & $vParam3 & ", " & $vParam4 & ")")
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_GET_BEST_ROW_ID
			If Not IsString($vParam1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
			If Not IsString($vParam2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
			If Not IsString($vParam3) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
			If Not __LO_IntIsBetween($vParam4, $LOB_DBASE_BEST_ROW_SCOPE_TEMPORARY, $LOB_DBASE_BEST_ROW_SCOPE_SESSION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)
			If Not IsBool($vParam5) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ", " & $vParam2 & ", " & $vParam3 & ", " & $vParam4 & ", " & $vParam5 & ")")
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_GET_INDEX_INFO
			If Not IsString($vParam1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
			If Not IsString($vParam2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
			If Not IsString($vParam3) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
			If Not IsBool($vParam4) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)
			If Not IsBool($vParam5) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ", " & $vParam2 & ", " & $vParam3 & ", " & $vParam4 & ", " & $vParam5 & ")")
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Case $LOB_DBASE_META_GET_CROSS_REF
			If Not IsString($vParam1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
			If Not IsString($vParam2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
			If Not IsString($vParam3) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
			If Not IsString($vParam4) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)
			If Not IsString($vParam5) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)
			If Not IsString($vParam6) Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

			$vReturn = Execute("$oConnection.MetaData" & $sCall & "(" & $vParam1 & ", " & $vParam2 & ", " & $vParam3 & ", " & $vParam4 & ", " & $vParam5 & ", " & $vParam6 & ")")
			If Not IsObj($vReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	EndSwitch

	Return SetError($__LO_STATUS_SUCCESS, 0, $vReturn)
EndFunc   ;==>_LOBase_DatabaseMetaDataQuery

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseName
; Description ...: Retrieve the Database Name property. See remarks.
; Syntax ........: _LOBase_DatabaseName(ByRef $oDBase)
; Parameters ....: $oDBase              - [in/out] an object.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDBase not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Database Source name.
;                  @Error 3 @Extended 2 Return 0 = Failed to convert URL to Computer Path.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning the Name value as a String. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the database data source is registered, then the Name property denotes the registration name. Otherwise, the name property contains the URL of the file.
;                  If the same database data source is registered under different names, the value of the Name property is not defined.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseName(ByRef $oDBase)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sName

	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sName = $oDBase.DatabaseDocument.DataSource.Name()
	If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If StringInStr($sName, "/") Then $sName = _LO_PathConvert($sName, $LO_PATHCONV_PCPATH_RETURN)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sName)
EndFunc   ;==>_LOBase_DatabaseName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRegisteredAdd
; Description ...: Add a Database to the Registered Databases list.
; Syntax ........: _LOBase_DatabaseRegisteredAdd(ByRef $oDBase, $sName)
; Parameters ....: $oDBase              - [in/out] an object. A Database object returned by a previous _LOBase_DatabaseGetObjByDoc or _LOBase_DatabaseGetObjByURL function.
;                  $sName               - a string value. The name to register the Database under.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDBase not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = Registered Database already exists with name as called in $sName.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to register Database.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Database successfully registered.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseRegisteredExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRegisteredAdd(ByRef $oDBase, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext

	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	If $oDBaseContext.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oDBaseContext.registerObject($sName, $oDBase)
	If Not $oDBaseContext.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DatabaseRegisteredAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRegisteredExists
; Description ...: Query whether a Registered Database exists by name.
; Syntax ........: _LOBase_DatabaseRegisteredExists($sName)
; Parameters ....: $sName               - a string value. The Database name to look for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sName not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query for Registered Database.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If Registered Database with called name exists, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRegisteredExists($sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext
	Local $bReturn

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$bReturn = $oDBaseContext.hasByName($sName)
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_DatabaseRegisteredExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRegisteredGetNames
; Description ...: Retrieve an array of Registered Database names.
; Syntax ........: _LOBase_DatabaseRegisteredGetNames()
; Parameters ....: None
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Registered Database Names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning array of Registered Database names. @Extended is set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseGetObjByURL
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRegisteredGetNames()
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext
	Local $asNames[0]

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$asNames = $oDBaseContext.getElementNames()
	If Not IsArray($asNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asNames), $asNames)
EndFunc   ;==>_LOBase_DatabaseRegisteredGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRegisteredRemoveByName
; Description ...: Remove a registered Database by name.
; Syntax ........: _LOBase_DatabaseRegisteredRemoveByName($sName)
; Parameters ....: $sName               - a string value. The Registered Database name to remove.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sName not a String.
;                  @Error 1 @Extended 2 Return 0 = No Registered Database found with name as called in $sName.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.sdb.DatabaseContext" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to unregister Database.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Database successfully unregistered.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_DatabaseRegisteredAdd, _LOBase_DatabaseRegisteredExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRegisteredRemoveByName($sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDBaseContext

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDBaseContext = $oServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
	If Not IsObj($oDBaseContext) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	If Not $oDBaseContext.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDBaseContext.revokeObject($sName)
	If $oDBaseContext.hasByName($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DatabaseRegisteredRemoveByName

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRequiresPassword
; Description ...: Check if a Database requires a password in order to connect to it.
; Syntax ........: _LOBase_DatabaseRequiresPassword(ByRef $oDBase)
; Parameters ....: $oDBase              - [in/out] an object. A Database object returned by a previous _LOBase_DatabaseGetObjByDoc or _LOBase_DatabaseGetObjByURL function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDBase not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to query Database for password requirement.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If Database requires a password to connect to it, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRequiresPassword(ByRef $oDBase)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn

	If Not IsObj($oDBase) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$bReturn = $oDBase.DatabaseDocument.DataSource.IsPasswordRequired()
	If Not IsBool($bReturn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)
EndFunc   ;==>_LOBase_DatabaseRequiresPassword

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DatabaseRollback
; Description ...: Rollback any changes made to the database since the last save.
; Syntax ........: _LOBase_DatabaseRollback(ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. A Connection object returned by a previous _LOBase_DatabaseConnectionGet function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oConnection not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oConnection not a connection Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully executed rollback command.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is unnecessary if AutoCommit is active (default).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DatabaseRollback(ByRef $oConnection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oConnection) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oConnection.supportsService("com.sun.star.sdbc.Connection") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oConnection.rollback()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_DatabaseRollback
