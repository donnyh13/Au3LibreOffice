#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Internal.au3"
#include "LibreOfficeWriter_Helper.au3"

; include in order as they were split from LibreOfficeWriter.au3
#include "LibreOfficeWriter_Doc.au3"
#include "LibreOfficeWriter_Frame.au3"
#include "LibreOfficeWriter_Table.au3"
#include "LibreOfficeWriter_DirectFormating.au3"
#include "LibreOfficeWriter_Field.au3"
#include "LibreOfficeWriter_Cell.au3"
#include "LibreOfficeWriter_FootEndNotes.au3"
#include "LibreOfficeWriter_Shapes.au3"
#include "LibreOfficeWriter_Images.au3"

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter)
; AutoIt Version : v3.3.16.1
; UDF Version    : 0.0.0.3
; Description ...: Provides basic functionality through Autoit for interacting with Libre Office Writer.
; Author(s) .....: donnyh13, mLipok
; Sources .......: jguinch -- Printmgr.au3, used (_PrintMgr_EnumPrinter);
;					mLipok -- OOoCalc.au3, used (__OOoCalc_ComErrorHandler_UserFunction,_InternalComErrorHandler,
;						-- WriterDemo.au3, used _CreateStruct;
;					Andrew Pitonyak & Laurent Godard (VersionGet);
;					Leagnus & GMK -- OOoCalc.au3, used (SetPropertyValue)
; Dll ...........:
; Note...........: Tips/templates taken from OOoCalc UDF written by user GMK; also from Word UDF by user water.
;					I found the book by Andrew Pitonyak very helpful also, titled, "OpenOffice.org Macros Explained;
;						OOME Third Edition".
;					Of course, this UDF is written using the English version of LibreOffice, and may only work for the English
;						version of LibreOffice installations. Many functions in this UDF may or may not work with OpenOffice
;						Writer, however some settings are definitely for LibreOffice only.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_SearchDescriptorCreate
; _LOWriter_SearchDescriptorModify
; _LOWriter_SearchDescriptorSimilarityModify
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_SearchDescriptorCreate
; Description ...: Create a Search Descriptor for searching a document.
; Syntax ........: _LOWriter_SearchDescriptorCreate(Byref $oDoc[, $bBackwards = False[, $bMatchCase = False[, $bWholeWord = False[, $bRegExp = False[, $bStyles = False[, $bSearchPropValues = False]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bBackwards          - [optional] a boolean value. Default is False. If True, the document is searched backwards.
;                  $bMatchCase          - [optional] a boolean value. Default is False. If True, the case of the letters is important for the Search.
;                  $bWholeWord          - [optional] a boolean value. Default is False. If True, only complete words will be found.
;                  $bRegExp             - [optional] a boolean value. Default is False. If True, the search string is evaluated as a regular expression.
;                  $bStyles             - [optional] a boolean value. Default is False. If True, the string is considered a Paragraph Style name, and the search will return any paragraph utilizing the specified name, EXCEPT if you input Format properties to search for, then setting this to True causes the search to search both for direct formatting matching those properties and also Paragraph/Character styles that contain matching properties.
;                  $bSearchPropValues   - [optional] a boolean value. Default is False. If True, any formatting properties searched for are matched based on their value, else if false, the search only looks for their existence. See Remarks.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bBackwards not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bMatchCase not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bWholeWord not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bRegExp not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $bStyles not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $bSearchPropValues not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create Search Descriptor.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returns a Search Descriptor Object for setting Search options.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bSearchPropValues is equivalent to the difference in selecting "Format" options in Libre Office's search box and "Attributes".
;				   Setting $bSearchPropValues to True, means that the search will look for matches using the specified property AND having the specified value, such as Character Weight, Bold, only matches that have Character weight of Bold will be returned, whereas if $bSearchPropValues is set to false, the search only looks for matches that have the specified property, regardless of its value. Such as Character weight, would match Bold, Semi-Bold, etc. From my understanding, the search is based on anything directly formatted unless $bStyles is also true.
;				   Note: The returned Search Descriptor is only good for the Document it was created by, it WILL NOT work for other documents.
; Related .......: _LOWriter_SearchDescriptorModify, _LOWriter_SearchDescriptorSimilarityModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_SearchDescriptorCreate(ByRef $oDoc, $bBackwards = False, $bMatchCase = False, $bWholeWord = False, $bRegExp = False, $bStyles = False, $bSearchPropValues = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSrchDescript

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If Not IsBool($bBackwards) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bMatchCase) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bWholeWord) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bRegExp) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bStyles) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bSearchPropValues) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)

	$oSrchDescript = $oDoc.createSearchDescriptor()
	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	With $oSrchDescript
		.SearchBackwards = $bBackwards
		.SearchCaseSensitive = $bMatchCase
		.SearchWords = $bWholeWord
		.SearchRegularExpression = $bRegExp
		.SearchStyles = $bStyles
		.ValueSearch = $bSearchPropValues
	EndWith

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oSrchDescript)
EndFunc   ;==>_LOWriter_SearchDescriptorCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_SearchDescriptorModify
; Description ...: Modify Search Descriptor settings of an existing Search Descriptor Object.
; Syntax ........: _LOWriter_SearchDescriptorModify(Byref $oSrchDescript[, $bBackwards = Null[, $bMatchCase = Null[, $bWholeWord = Null[, $bRegExp = Null[, $bStyles = Null[, $bSearchPropValues = Null]]]]]])
; Parameters ....: $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $bBackwards          - [optional] a boolean value. Default is False. If True, the document is searched backwards.
;                  $bMatchCase          - [optional] a boolean value. Default is False. If True, the case of the letters is important for the Search.
;                  $bWholeWord          - [optional] a boolean value. Default is False. If True, only complete words will be found.
;                  $bRegExp             - [optional] a boolean value. Default is False. If True, the search string is evaluated as a regular expression. Cannot be set to True if Similarity Search is set to True.
;                  $bStyles             - [optional] a boolean value. Default is False. If True, the string is considered a Paragraph Style name, and the search will return any paragraph utilizing the specified name, EXCEPT if you input Format properties to search for, then setting this to True causes the search to search both for direct formatting matching those properties and also Paragraph/Character styles that contain matching properties.
;                  $bSearchPropValues   - [optional] a boolean value. Default is False. If True, any formatting properties searched for are matched based on their value, else if false, the search only looks for their existence. See Remarks.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript Object not a Search Descriptor Object.
;				   @Error 1 @Extended 3 Return 0 = $bBackwards not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bMatchCase not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bWholeWord not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $bRegExp not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $bStyles not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $bSearchPropValues not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = $bRegExp is set to True while Similarity Search is also set to True.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Returns 1 after directly modifying Search Descriptor Object.
; ;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bSearchPropValues is equivalent to the difference in selecting "Format" options in Libre Office's search box and "Attributes". Setting $bSearchPropValues to True, means that the search will look for matches using the specified property AND having the specified value, such as Character Weight, Bold, only matches that have Character weight of Bold will be returned, whereas if $bSearchPropValues is set to false, the search only looks for matches that have the specified property, regardless of its value. Such as Character weight, would match Bold, Semi-Bold, etc. From my understanding, the search is based on anything directly formatted unless $bStyles is also true.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_SearchDescriptorCreate, _LOWriter_SearchDescriptorSimilarityModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_SearchDescriptorModify(ByRef $oSrchDescript, $bBackwards = Null, $bMatchCase = Null, $bWholeWord = Null, $bRegExp = Null, $bStyles = Null, $bSearchPropValues = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSrchDescript[6]

	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bBackwards, $bMatchCase, $bWholeWord, $bRegExp, $bStyles, $bSearchPropValues) Then
		__LOWriter_ArrayFill($avSrchDescript, $oSrchDescript.SearchBackwards(), $oSrchDescript.SearchCaseSensitive(), $oSrchDescript.SearchWords(), _
				$oSrchDescript.SearchRegularExpression(), $oSrchDescript.SearchStyles(), $oSrchDescript.getValueSearch())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avSrchDescript)
	EndIf

	If ($bBackwards <> Null) Then
		If Not IsBool($bBackwards) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$oSrchDescript.SearchBackwards = $bBackwards
	EndIf

	If ($bMatchCase <> Null) Then
		If Not IsBool($bMatchCase) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oSrchDescript.SearchCaseSensitive = $bMatchCase
	EndIf

	If ($bWholeWord <> Null) Then
		If Not IsBool($bWholeWord) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oSrchDescript.SearchWords = $bWholeWord
	EndIf

	If ($bRegExp <> Null) Then
		If Not IsBool($bRegExp) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If ($bRegExp = True) And ($oSrchDescript.SearchSimilarity = True) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		$oSrchDescript.SearchRegularExpression = $bRegExp
	EndIf

	If ($bStyles <> Null) Then
		If Not IsBool($bStyles) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oSrchDescript.SearchStyles = $bStyles
	EndIf

	If ($bSearchPropValues <> Null) Then
		If Not IsBool($bSearchPropValues) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oSrchDescript.ValueSearch = $bSearchPropValues
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_SearchDescriptorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_SearchDescriptorSimilarityModify
; Description ...: Modify Similarity Search Settings for an existing Search Descriptor Object.
; Syntax ........: _LOWriter_SearchDescriptorSimilarityModify(Byref $oSrchDescript[, $bSimilarity = Null[, $bCombine = Null[, $iRemove = Null[, $iAdd = Null[, $iExchange = Null]]]]])
; Parameters ....: $oSrchDescript       - [in/out] an object. A Search Descriptor Object returned from _LOWriter_SearchDescriptorCreate function.
;                  $bSimilarity         - [optional] a boolean value. Default is Null. If True, a "similarity search" is performed.
;                  $bCombine            - [optional] a boolean value. Default is Null. If True, all similarity rules ($iRemove, $iAdd, and $iExchange) are applied together.
;                  $iRemove             - [optional] an integer value. Default is Null. Specifies the number of characters that may be ignored to match the search pattern.
;                  $iAdd                - [optional] an integer value. Default is Null. Specifies the number of characters that must be added to match the search pattern.
;                  $iExchange           - [optional] an integer value. Default is Null. Specifies the number of characters that must be replaced to match the search pattern.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSrchDescript not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSrchDescript Object not a Search Descriptor Object.
;				   @Error 1 @Extended 3 Return 0 = $bSimilarity not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bCombine not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iRemove, $iAdd, or $iExchange set to a value, but $bSimilarity not set to True.
;				   @Error 1 @Extended 6 Return 0 = $iRemove not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iAdd not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iExchange not an Integer.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = $bSimilarity is set to True while Regular Expression Search is also set to True.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Returns 1 after directly modifying Search Descriptor Object.
; ;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_SearchDescriptorCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_SearchDescriptorSimilarityModify(ByRef $oSrchDescript, $bSimilarity = Null, $bCombine = Null, $iRemove = Null, $iAdd = Null, $iExchange = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avSrchDescript[5]

	If Not IsObj($oSrchDescript) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not $oSrchDescript.supportsService("com.sun.star.util.SearchDescriptor") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bSimilarity, $bCombine, $iRemove, $iAdd, $iExchange) Then
		__LOWriter_ArrayFill($avSrchDescript, $oSrchDescript.SearchSimilarity(), $oSrchDescript.SearchSimilarityRelax(), _
				$oSrchDescript.SearchSimilarityRemove(), $oSrchDescript.SearchSimilarityAdd(), $oSrchDescript.SearchSimilarityExchange())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avSrchDescript)
	EndIf

	If ($bSimilarity <> Null) Then
		If Not IsBool($bSimilarity) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		If ($bSimilarity = True) And ($oSrchDescript.SearchRegularExpression = True) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		$oSrchDescript.SearchSimilarity = $bSimilarity
	EndIf

	If ($bCombine <> Null) Then
		If Not IsBool($bCombine) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oSrchDescript.SearchSimilarityRelax = $bCombine
	EndIf

	If Not __LOWriter_VarsAreNull($iRemove, $iAdd, $iExchange) Then
		If ($oSrchDescript.SearchSimilarity() = False) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		If ($iRemove <> Null) Then
			If Not IsInt($iRemove) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			$oSrchDescript.SearchSimilarityRemove = $iRemove
		EndIf

		If ($iAdd <> Null) Then
			If Not IsInt($iAdd) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
			$oSrchDescript.SearchSimilarityAdd = $iAdd
		EndIf

		If ($iExchange <> Null) Then
			If Not IsInt($iExchange) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
			$oSrchDescript.SearchSimilarityExchange = $iExchange
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_SearchDescriptorSimilarityModify

