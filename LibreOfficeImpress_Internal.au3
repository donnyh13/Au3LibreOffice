#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Impress
#include "LibreOfficeImpress_Constants.au3"
#include "LibreOfficeImpress_Helper.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Various functions for internal data processing, data retrieval, retrieving and applying settings for LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __LOImpress_AddTo1DArray
; __LOImpress_ArrayFill
; __LOImpress_CreateStruct
; __LOImpress_DrawShape_GetCustomType
; __LOImpress_DrawShapeGetType
; __LOImpress_FilterNameGet
; __LOImpress_GradientNameInsert
; __LOImpress_GradientPresets
; __LOImpress_InternalComErrorHandler
; __LOImpress_IntIsBetween
; __LOImpress_NumIsBetween
; __LOImpress_SetPropertyValue
; __LOImpress_TransparencyGradientConvert
; __LOImpress_TransparencyGradientNameInsert
; __LOImpress_UnitConvert
; __LOImpress_VarsAreNull
; __LOImpress_VersionCheck
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_AddTo1DArray
; Description ...: Add data to a 1 Dimensional array.
; Syntax ........: __LOImpress_AddTo1DArray(ByRef $aArray, $vData[, $bCountInFirst = False])
; Parameters ....: $aArray              - [in/out] an array of unknowns. The Array to directly add data to. Array will be directly modified.
;                  $vData               - a variant value. The Data to add to the Array.
;                  $bCountInFirst       - [optional] a boolean value. Default is False. If True the first element of the array is a count of contained elements.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $aArray not an Array
;                  @Error 1 @Extended 2 Return 0 = $bCountinFirst not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $aArray contains too many columns.
;                  @Error 1 @Extended 4 Return 0 = $aArray[0] contains non integer data or is not empty, and $bCountInFirst is set to True.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Array item was successfully added.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_AddTo1DArray(ByRef $aArray, $vData, $bCountInFirst = False)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($aArray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bCountInFirst) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If UBound($aArray, $UBOUND_COLUMNS) > 1 Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Too many columns

	If $bCountInFirst And (UBound($aArray) = 0) Then
		ReDim $aArray[1]
		$aArray[0] = 0
	EndIf

	If $bCountInFirst And (($aArray[0] <> "") And Not IsInt($aArray[0])) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	ReDim $aArray[UBound($aArray) + 1]
	$aArray[UBound($aArray) - 1] = $vData
	If $bCountInFirst Then $aArray[0] += 1

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_AddTo1DArray

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_ArrayFill
; Description ...: Fill an Array with data.
; Syntax ........: __LOImpress_ArrayFill(ByRef $aArrayToFill[, $vVar1 = Null[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null[, $vVar9 = Null[, $vVar10 = Null[, $vVar11 = Null[, $vVar12 = Null[, $vVar13 = Null[, $vVar14 = Null[, $vVar15 = Null[, $vVar16 = Null[, $vVar17 = Null[, $vVar18 = Null]]]]]]]]]]]]]]]]]])
; Parameters ....: $aArrayToFill        - [in/out] an array of unknowns. The Array to Fill. Array will be directly modified.
;                  $vVar1               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar2               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar3               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar4               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar5               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar6               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar7               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar8               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar9               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar10              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar11              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar12              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar13              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar14              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar15              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar16              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar17              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar18              - [optional] a variant value. Default is Null. The Data to add to the Array.
; Return values .: None
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call only how many you parameters you need to add to the Array. Automatically resizes the Array if it is the incorrect size.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_ArrayFill(ByRef $aArrayToFill, $vVar1 = Null, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null, $vVar9 = Null, $vVar10 = Null, $vVar11 = Null, $vVar12 = Null, $vVar13 = Null, $vVar14 = Null, $vVar15 = Null, $vVar16 = Null, $vVar17 = Null, $vVar18 = Null)
	#forceref $vVar1, $vVar2, $vVar3, $vVar4, $vVar5, $vVar6, $vVar7, $vVar8, $vVar9, $vVar10, $vVar11, $vVar12, $vVar13, $vVar14, $vVar15, $vVar16, $vVar17, $vVar18

	If UBound($aArrayToFill) < (@NumParams - 1) Then ReDim $aArrayToFill[@NumParams - 1]
	For $i = 0 To @NumParams - 2
		$aArrayToFill[$i] = Eval("vVar" & $i + 1)
	Next
EndFunc   ;==>__LOImpress_ArrayFill

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_CreateStruct
; Description ...: Creates a Struct.
; Syntax ........: __LOImpress_CreateStruct($sStructName)
; Parameters ....: $sStructName         - a string value. Name of structure to create.
; Return values .: Success: Structure.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sStructName not a string
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object
;                  @Error 2 @Extended 2 Return 0 = Error creating requested structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Structure = Success. Property Structure Returned
; Author ........: mLipok
; Modified ......: donnyh13 - Added error checking.
; Remarks .......: From WriterDemo.au3 as modified by mLipok from WriterDemo.vbs found in the LibreOffice SDK examples.
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/204665-libreopenoffice-writer/?do=findComment&comment=1471711
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_CreateStruct($sStructName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $tStruct

	If Not IsString($sStructName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tStruct = $oServiceManager.Bridge_GetStruct($sStructName)
	If Not IsObj($tStruct) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $tStruct)
EndFunc   ;==>__LOImpress_CreateStruct

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShape_GetCustomType
; Description ...: Return the Shape Type Constant corresponding to the Custom Shape Type string.
; Syntax ........: __LOImpress_DrawShape_GetCustomType($sCusShapeType)
; Parameters ....: $sCusShapeType       - a string value. The Returned Custom Shape Type Value from CustomShapeGeometry Array of properties.
; Return values .: Success: Integer or -1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sCusShapeType not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Custom Shape Type was successfully identified. Returning the Constant value of the Shape, see Constants $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 0 @Extended 0 Return -1 = Success. Custom Shape is of an unimplemented type that has an ambiguous name, and cannot be identified. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Some shapes are not implemented, or not fully implemented into LibreOffice for automation, consequently they do not have appropriate type names as of yet. Many have simply ambiguous names, such as "non-primitive".
;                  #1 Because of this the following shape types cannot be identified, and this function will return -1:
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT, known as "mso-spt100".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, known as "non-primitive", should be "corner-right-arrow".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT, known as "non-primitive", should be "right-left-arrow".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED, known as "non-primitive", should be "s-sharped-arrow".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT, known as "non-primitive", should be "split-arrow".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT, known as "mso-spt100", should be "striped-right-arrow".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT, known as "mso-spt89", should be "up-right-arrow-callout".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, known as "mso-spt100", should be "up-right-down-arrow".
;                   $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE, known as "mso-spt100", should be "circle-pie".
;                   $LOI_DRAWSHAPE_TYPE_STARS_6_POINT, known as "non-primitive", should be "star6".
;                   $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE, known as "non-primitive", should be "concave-star6".
;                   $LOI_DRAWSHAPE_TYPE_STARS_12_POINT, known as "non-primitive", should be "star12".
;                   $LOI_DRAWSHAPE_TYPE_STARS_SIGNET, known as "non-primitive", should be "signet".
;                   $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD, known as "non-primitive", should be "cloud"?
;                   $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER, known as "non-primitive", should be "flower"?
;                   $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING, known as "non-primitive", should be "lightning".
;                  #2 The following Shapes implement the same type names, and are consequently indistinguishable:
;                   $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE, $LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE is returned for either one.)
;                   $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE is returned for either one.)
;                   $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED is returned for either one.)
;                  #3 The following Shapes have strange names that may change in the future, but currently are able to be identified:
;                   $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE, known as, "mso-spt21", should be "doorplate"
;                   $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, known as, "col-502ad400", should be ??
;                   $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, known as, "col-60da8460", should be ??
;                  #4 The following Shapes are customizable one to another, and are consequently indistinguishable:
;                   $LOI_DRAWSHAPE_TYPE_FONTWORK_* (The Value of $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL is returned for any of these.)
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShape_GetCustomType($sCusShapeType)
	If Not IsString($sCusShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $sCusShapeType
		Case "quad-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_4_WAY)

		Case "quad-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY)

		Case "down-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN)

		Case "left-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT)

		Case "left-right-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT)

		Case "right-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT)

		Case "up-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP)

		Case "up-down-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN)

			;~ 	Case "mso-spt100" ; Can't include this one as other shapes return mso-spt100 also
			;~ Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT)

		Case "circular-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CIRCULAR)

		Case "corner-right-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT)

		Case "down-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_DOWN)

		Case "left-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT)

		Case "left-right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT)

		Case "notched-right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT)

		Case "right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT)

		Case "right-left-arrow" ; "non-primitive"??

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT)

		Case "s-sharped-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED)

		Case "split-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT)

		Case "striped-right-arrow" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT)

		Case "up-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP)

		Case "up-down-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_DOWN)

		Case "up-right-arrow-callout", "mso-spt89" ; "mso-spt89"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT)

		Case "up-right-down-arrow" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN)

		Case "chevron"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_CHEVRON)

		Case "pentagon-right"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_ARROWS_PENTAGON)

		Case "block-arc"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_ARC_BLOCK)

		Case "circle-pie" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE)

		Case "cross"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CROSS)

		Case "cube"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CUBE)

		Case "can"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CYLINDER)

		Case "diamond"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_DIAMOND)

		Case "ellipse"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE)
			;~ $LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE

		Case "paper"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_FOLDED_CORNER)

		Case "frame" ; Not working

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_FRAME)

		Case "hexagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_HEXAGON)

		Case "octagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_OCTAGON)

		Case "parallelogram"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_PARALLELOGRAM)

		Case "rectangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE)
			;~ $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE

		Case "round-rectangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED)
			;~ $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED

		Case "pentagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_REGULAR_PENTAGON)

		Case "ring"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_RING)

		Case "trapezoid"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_TRAPEZOID)

		Case "isosceles-triangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES)

		Case "right-triangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_RIGHT)

		Case "cloud-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_CLOUD)

		Case "line-callout-1"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_1)

		Case "line-callout-2"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_2)

		Case "line-callout-3"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_3)

		Case "rectangular-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR)

		Case "round-rectangular-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED)

		Case "round-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_CALLOUT_ROUND)

		Case "flowchart-card"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_CARD)

		Case "flowchart-collate"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_COLLATE)

		Case "flowchart-connector"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR)

		Case "flowchart-off-page-connector"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE)

		Case "flowchart-data"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DATA)

		Case "flowchart-decision"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DECISION)

		Case "flowchart-delay"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DELAY)

		Case "flowchart-direct-access-storage"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE)

		Case "flowchart-display"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DISPLAY)

		Case "flowchart-document"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_DOCUMENT)

		Case "flowchart-extract"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_EXTRACT)

		Case "flowchart-internal-storage"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE)

		Case "flowchart-magnetic-disk"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MAGNETIC_DISC)

		Case "flowchart-manual-input"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_INPUT)

		Case "flowchart-manual-operation"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_OPERATION)

		Case "flowchart-merge"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MERGE)

		Case "flowchart-multidocument"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_MULTIDOCUMENT)

		Case "flowchart-or"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_OR)

		Case "flowchart-preparation"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PREPARATION)

		Case "flowchart-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS)

		Case "flowchart-alternate-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE)

		Case "flowchart-predefined-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED)

		Case "flowchart-punched-tape"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_PUNCHED_TAPE)

		Case "flowchart-sequential-access"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS)

		Case "flowchart-sort"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_SORT)

		Case "flowchart-stored-data"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_STORED_DATA)

		Case "flowchart-summing-junction"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION)

		Case "flowchart-terminator"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FLOWCHART_TERMINATOR)

		Case "fontwork-arch-down-pour", "fontwork-arch-left-pour", "fontwork-arch-right-pour", "fontwork-arch-up-pour", "fontwork-arch-down-curve", _
				"fontwork-arch-left-curve", "fontwork-arch-right-curve", "fontwork-arch-up-curve", "fontwork-chevron-down", "fontwork-chevron-up", _
				"fontwork-circle-curve", "fontwork-circle-pour", "fontwork-curve-down", "fontwork-curve-up", "fontwork-fade-down", "fontwork-fade-left", _
				"fontwork-fade-right", "fontwork-fade-up", "fontwork-fade-up-and-left", "fontwork-fade-up-and-right", "fontwork-inflate", "fontwork-slant-down", _
				"fontwork-slant-up", "fontwork-stop", "fontwork-triangle-up", "fontwork-triangle-down", "fontwork-open-circle-curve", "fontwork-open-circle-pour", _
				"fontwork-plain-text", "fontwork-wave"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL) ; Can't differentiate reliably

		Case "star4"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_4_POINT)

		Case "star5"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_5_POINT)

		Case "star6" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_6_POINT)

		Case "concave-star6" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE)

		Case "star8"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_8_POINT)

		Case "star12" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_12_POINT)

		Case "star24"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_24_POINT)

		Case "mso-spt21", "doorplate" ; "doorplate"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE)

		Case "bang"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_EXPLOSION)

		Case "horizontal-scroll"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_SCROLL_HORIZONTAL)

		Case "vertical-scroll"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_SCROLL_VERTICAL)

		Case "signet" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_STARS_SIGNET)

		Case "col-502ad400"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND)

		Case "col-60da8460"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON)

		Case "quad-bevel"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_SQUARE)

		Case "brace-pair"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_DOUBLE)

		Case "left-brace"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_LEFT)

		Case "right-brace"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_RIGHT)

		Case "bracket-pair"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_DOUBLE)

		Case "left-bracket"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_LEFT)

		Case "right-bracket"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_RIGHT)

		Case "cloud"
			;~ Custom Shape Geometry Type = "non-primitive" ???? Try "cloud"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD)

		Case "flower"
			;~ Custom Shape Geometry Type = "non-primitive" ???? Try "flower"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER)

		Case "heart"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_HEART)

		Case "lightning"
			;~ Custom Shape Geometry Type = "non-primitive" ???? Try "lightning"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING)

		Case "moon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_MOON)

		Case "smiley"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_SMILEY)

		Case "sun"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_SUN)

		Case "forbidden"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_PROHIBITED)

		Case "puzzle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE)

		Case Else

			Return SetError($__LO_STATUS_SUCCESS, 0, -1)
	EndSwitch
EndFunc   ;==>__LOImpress_DrawShape_GetCustomType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_DrawShapeGetType
; Description ...: Identify a Shape's type.
; Syntax ........: __LOImpress_DrawShapeGetType(ByRef $oShape)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_ShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shape's type.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify Shape.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning the Shape's Type, corresponding to one of the Constants $LOI_SHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_DrawShapeGetType(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avShapeTypes[20][2] = [[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.Shape3DSceneObject"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.CustomShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.MeasureShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.EllipseShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.ClosedBezierShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.OpenBezierShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.PolyPolygonShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.PolyLineShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.LineShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.ConnectorShape"], [$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.OpenFreeHandShape"], _
			[$LOI_SHAPE_TYPE_DRAWING_SHAPE, "com.sun.star.drawing.ClosedFreeHandShape"], [$LOI_SHAPE_TYPE_FORM_CONTROL, "com.sun.star.drawing.ControlShape"], _
			[$LOI_SHAPE_TYPE_IMAGE, "com.sun.star.drawing.GraphicObjectShape"], [$LOI_SHAPE_TYPE_MEDIA, "com.sun.star.drawing.MediaShape"], _
			[$LOI_SHAPE_TYPE_OLE2, "com.sun.star.drawing.OLE2Shape"], [$LOI_SHAPE_TYPE_TABLE, "com.sun.star.drawing.TableShape"], _
			[$LOI_SHAPE_TYPE_TEXTBOX, "com.sun.star.drawing.TextShape"], [$LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE, "com.sun.star.presentation.SubtitleShape"], _
			[$LOI_SHAPE_TYPE_TEXTBOX_TITLE, "com.sun.star.presentation.TitleTextShape"]]
	Local $sShapeType

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sShapeType = $oShape.ShapeType()
	If Not IsString($sShapeType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($avShapeTypes) - 1
		If ($sShapeType = $avShapeTypes[$i][1]) Then Return SetError($__LO_STATUS_SUCCESS, 0, $avShapeTypes[$i][0])
	Next

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
EndFunc   ;==>__LOImpress_DrawShapeGetType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_FilterNameGet
; Description ...: Retrieves the correct L.O. Filter name for use in SaveAs and Export.
; Syntax ........: __LOImpress_FilterNameGet(ByRef $sDocSavePath[, $bExportFilters = False])
; Parameters ....: $sDocSavePath        - [in/out] a string value. Full path with extension.
;                  $bExportFilters      - [optional] a boolean value. Default is False. If True, includes the Filter Names that can be used to Export only, in the search.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sDocSavePath is not a string.
;                  @Error 1 @Extended 2 Return 0 = $bExportFilters not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sDocSavePath is not a correct path or URL.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Success. Returns required filter name from "SaveAs" Filter Names.
;                  @Error 0 @Extended 2 Return String = Success. Returns required filter name from "Export" Filter Names.
;                  @Error 0 @Extended 3 Return String = Filter Name not found for given file extension, defaulting to .odp file format and updating save path accordingly.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Searches a predefined list of extensions stored in an array. Not all FilterNames are listed.
;                  For finding your own Filter Names, see convertfilters.html found in L.O. Install Folder: LibreOffice\help\en-US\text\shared\guide
;                  Or See: "OOME_3_0", "OpenOffice.org Macros Explained OOME Third Edition" by Andrew D. Pitonyak, which has a handy Macro for listing all Filter Names, found on page 284 of the above book in the ODT format.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_FilterNameGet(ByRef $sDocSavePath, $bExportFilters = False)
	Local $iLength, $iSlashLocation, $iDotLocation
	Local Const $STR_NOCASESENSE = 0, $STR_STRIPALL = 8
	Local $sFileExtension, $sFilterName
	Local $msSaveAsFilters[], $msExportFilters[]

	If Not IsString($sDocSavePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExportFilters) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iLength = StringLen($sDocSavePath)

	$msSaveAsFilters[".fodp"] = "OpenDocument Presentation Flat XML"
	$msSaveAsFilters[".pot"] = "PowerPoint 3"
	$msSaveAsFilters[".potx"] = "Impress MS PowerPoint 2007 XML Template" ; Note these have a XML version too.
	$msSaveAsFilters[".pps"] = "MS PowerPoint 97 AutoPlay"
	$msSaveAsFilters[".ppt"] = "PowerPoint 3"
	$msSaveAsFilters[".ppsx"] = "Impress MS PowerPoint 2007 XML AutoPlay" ; Note these have a XML version too.
	$msSaveAsFilters[".pptm"] = "Impress MS PowerPoint 2007 XML VBA"
	$msSaveAsFilters[".pptx"] = "Impress MS PowerPoint 2007 XML" ; Note these have a XML version too.
	$msSaveAsFilters[".odg"] = "impress8_draw"
	$msSaveAsFilters[".odp"] = "impress8"
	$msSaveAsFilters[".otp"] = "impress8_template"
	$msSaveAsFilters[".uop"] = "UOF presentation"

	If $bExportFilters Then
		$msExportFilters[".apng"] = "impress_png_Export"
		$msExportFilters[".bmp"] = "impress_bmp_Export"
		$msExportFilters[".emf"] = "impress_emf_Export"
		$msExportFilters[".eps"] = "impress_eps_Export"
		$msExportFilters[".gif"] = "impress_gif_Export"
		$msExportFilters[".htm"] = "impress_html_Export"
		$msExportFilters[".html"] = "impress_html_Export"
		$msExportFilters[".jfif"] = "impress_jpg_Export"
		$msExportFilters[".jif"] = "impress_jpg_Export"
		$msExportFilters[".jpg"] = "impress_jpg_Export"
		$msExportFilters[".jpeg"] = "impress_jpg_Export"
		$msExportFilters[".svg"] = "impress_svg_Export"
		$msExportFilters[".pdf"] = "impress_pdf_Export"
		$msExportFilters[".png"] = "impress_png_Export"
		$msExportFilters[".tif"] = "impress_tif_Export"
		$msExportFilters[".tiff"] = "impress_tif_Export"
		$msExportFilters[".webp"] = "writer_webp_Export"
		$msExportFilters[".wmf"] = "impress_wmf_Export"
		$msExportFilters[".xhtml"] = "XHTML Impress File"
	EndIf

	If StringInStr($sDocSavePath, "file:///") Then ;  If L.O. URl Then
		$iSlashLocation = StringInStr($sDocSavePath, "/", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, ($iLength - $iDotLocation) + 1)

	ElseIf StringInStr($sDocSavePath, "\") Then ;  Else if PC Path Then
		$iSlashLocation = StringInStr($sDocSavePath, "\", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, $iLength - $iDotLocation + 1)

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	If $sFileExtension = $sDocSavePath Then ;  If no file extension identified, append .ods extension and return.
		$sDocSavePath = $sDocSavePath & ".odp"

		Return SetError($__LO_STATUS_SUCCESS, 3, "impress8")

	Else
		$sFileExtension = StringLower(StringStripWS($sFileExtension, $STR_STRIPALL))
	EndIf

	$sFilterName = $msSaveAsFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sFilterName)

	If $bExportFilters Then $sFilterName = $msExportFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 2, $sFilterName)

	$sDocSavePath = StringReplace($sDocSavePath, $sFileExtension, ".odp") ; If No results, replace with ODS extension.

	Return SetError($__LO_STATUS_SUCCESS, 3, "impress8")
EndFunc   ;==>__LOImpress_FilterNameGet

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_GradientNameInsert
; Description ...: Create and insert a new Gradient name.
; Syntax ........: __LOImpress_GradientNameInsert(ByRef $oDoc, $tGradient[, $sGradientName = "Gradient "])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $tGradient           - a dll struct value. A Gradient Structure to copy settings from.
;                  $sGradientName       - [optional] a string value. Default is "Gradient ". The Gradient name to create.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tGradient not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sGradientName not a string.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.GradientTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" or "com.sun.star.awt.Gradient2" structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error creating Gradient Name.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. A new Gradient name was created. Returning the new name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If The Gradient name is blank, I need to create a new name and apply it. I think I could re-use an old one without problems, but I'm not sure, so to be safe, I will create a new one.
;                  If there are no names that have been already created, then I need to create and apply one before the gradient will be displayed.
;                  Else if a preset Gradient is called, I need to create its name before it can be used.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_GradientNameInsert(ByRef $oDoc, $tGradient, $sGradientName = "Gradient ")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tNewGradient
	Local $oGradTable
	Local $iCount = 1
	Local $sGradient = "com.sun.star.awt.Gradient2"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If Not __LOImpress_VersionCheck(7.6) Then $sGradient = "com.sun.star.awt.Gradient"

	$oGradTable = $oDoc.createInstance("com.sun.star.drawing.GradientTable")
	If Not IsObj($oGradTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sGradientName = "Gradient ") Then
		While $oGradTable.hasByName($sGradientName & $iCount)
			$iCount += 1
			Sleep((IsInt($iCount / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
		WEnd
		$sGradientName = $sGradientName & $iCount
	EndIf

	$tNewGradient = __LOImpress_CreateStruct($sGradient)
	If Not IsObj($tNewGradient) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; Copy the settings over from the input Style Gradient to my new one. This may not be necessary? But just in case.
	With $tNewGradient
		.Style = $tGradient.Style()
		.XOffset = $tGradient.XOffset()
		.YOffset = $tGradient.YOffset()
		.Angle = $tGradient.Angle()
		.Border = $tGradient.Border()
		.StartColor = $tGradient.StartColor()
		.EndColor = $tGradient.EndColor()
		.StartIntensity = $tGradient.StartIntensity()
		.EndIntensity = $tGradient.EndIntensity()

		If __LOImpress_VersionCheck(7.6) Then .ColorStops = $tGradient.ColorStops()

	EndWith

	If Not $oGradTable.hasByName($sGradientName) Then
		$oGradTable.insertByName($sGradientName, $tNewGradient)
		If Not ($oGradTable.hasByName($sGradientName)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $sGradientName)
EndFunc   ;==>__LOImpress_GradientNameInsert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_GradientPresets
; Description ...: Set Page background Gradient to preset settings.
; Syntax ........: __LOImpress_GradientPresets(ByRef $oDoc, ByRef $oObject, ByRef $tGradient, $sGradientName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oObject             - [in/out] an object. The Object to modify the Gradient settings for.
;                  $tGradient           - [in/out] an object. The Fill Gradient Object to modify the Gradient settings for.
;                  $sGradientName       - a string value. The Gradient Preset name to apply.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.awt.ColorStop" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to create Gradient name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. The Style Gradient settings were successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_GradientPresets(ByRef $oDoc, ByRef $oObject, ByRef $tGradient, $sGradientName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tColorStop, $tStopColor
	Local $atColorStop[2]

	If __LOImpress_VersionCheck(7.6) Then
		$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$atColorStop[0] = $tColorStop

		$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$atColorStop[1] = $tColorStop
	EndIf

	Switch $sGradientName
		Case $LOI_GRAD_NAME_PASTEL_BOUQUET
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 14543051
				.EndColor = 16766935
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_PASTEL_DREAM
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RECT
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 450
				.Border = 0
				.StartColor = 16766935
				.EndColor = 11847644
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_BLUE_TOUCH
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 10
				.Border = 0
				.StartColor = 11847644
				.EndColor = 14608111
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_BLANK_W_GRAY
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 900
				.Border = 75
				.StartColor = $LOI_COLOR_WHITE
				.EndColor = 14540253
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_LONDON_MIST
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 13421772
				.EndColor = 6710886
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUBMARINE
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 14543051
				.EndColor = 11847644
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_MIDNIGHT
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = $LOI_COLOR_BLACK
				.EndColor = 2777241
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_DEEP_OCEAN
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 0
				.Border = 0
				.StartColor = $LOI_COLOR_BLACK
				.EndColor = 7512015
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_MAHOGANY
			With $tGradient
				.Style = $LOI_GRAD_TYPE_SQUARE
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 450
				.Border = 0
				.StartColor = $LOI_COLOR_BLACK
				.EndColor = 9250846
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_GREEN_GRASS
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 16776960
				.EndColor = 8508442
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_NEON_LIGHT
			With $tGradient
				.Style = $LOI_GRAD_TYPE_ELLIPTICAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 0
				.Border = 15
				.StartColor = 1209890
				.EndColor = $LOI_COLOR_WHITE
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNSHINE
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 66
				.YOffset = 33
				.Angle = 0
				.Border = 33
				.StartColor = 16760576
				.EndColor = 16776960
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_RAINBOW
			With $tGradient
				.Style = $LOI_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 100
				.Angle = 0
				.Border = 0
				.StartColor = $LOI_COLOR_WHITE
				.EndColor = $LOI_COLOR_WHITE
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					ReDim $atColorStop[7]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0.2
					$atColorStop[0] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.2

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 0
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.4

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 1
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0
					$tStopColor.Green = 1
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[3] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.65

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0
					$tStopColor.Green = 1
					$tStopColor.Blue = 1
					$tColorStop.StopColor = $tStopColor

					$atColorStop[4] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.8

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 0
					$tStopColor.Blue = 1
					$tColorStop.StopColor = $tStopColor

					$atColorStop[5] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.8
					$atColorStop[6] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNRISE
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 3713206
				.EndColor = 14065797
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					ReDim $atColorStop[4]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.505882352941176
					$tStopColor.Green = 0.784313725490196
					$tStopColor.Blue = 0.768627450980392
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.75

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.717647058823529
					$tStopColor.Green = 0.807843137254902
					$tStopColor.Blue = 0.698039215686275
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 1
					$atColorStop[3] = $tColorStop
				EndIf
			EndWith

		Case $LOI_GRAD_NAME_SUNDOWN
			With $tGradient
				.Style = $LOI_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 985943
				.EndColor = 16759664
				.StartIntensity = 100
				.EndIntensity = 100

				If __LOImpress_VersionCheck(7.6) Then
					ReDim $atColorStop[5]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.3

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.392156862745098
					$tStopColor.Green = 0.305882352941177
					$tStopColor.Blue = 0.690196078431373
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.827450980392157
					$tStopColor.Green = 0.572549019607843
					$tStopColor.Blue = 0.83921568627451
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.75

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.996078431372549
					$tStopColor.Green = 0.733333333333333
					$tStopColor.Blue = 0.76078431372549
					$tColorStop.StopColor = $tStopColor

					$atColorStop[3] = $tColorStop

					$tColorStop = __LOImpress_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 1
					$atColorStop[4] = $tColorStop
				EndIf
			EndWith

		Case Else ; Custom Gradient Name
			__LOImpress_GradientNameInsert($oDoc, $tGradient, $sGradientName)
			If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$oObject.FillGradientName = $sGradientName

			Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndSwitch

	If __LOImpress_VersionCheck(7.6) Then
		$tColorStop = $atColorStop[0] ; "Start Color" Value.

		$tStopColor = $tColorStop.StopColor()

		$tStopColor.Red = (BitAND(BitShift($tGradient.StartColor(), 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($tGradient.StartColor(), 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($tGradient.StartColor(), 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStop[0] = $tColorStop

		$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; Last element is "End Color" Value.

		$tStopColor = $tColorStop.StopColor()

		$tStopColor.Red = (BitAND(BitShift($tGradient.EndColor(), 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($tGradient.EndColor(), 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($tGradient.EndColor(), 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStop[UBound($atColorStop) - 1] = $tColorStop

		$tGradient.ColorStops = $atColorStop
	EndIf

	__LOImpress_GradientNameInsert($oDoc, $tGradient, $sGradientName)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oObject.FillGradient = $tGradient
	$oObject.FillGradientName = $sGradientName
	$oObject.FillGradientStepCount = $tGradient.StepCount()

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOImpress_GradientPresets

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LOImpress_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _LOImpress_ComError_UserFunction(Default)
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
EndFunc   ;==>__LOImpress_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_IntIsBetween
; Description ...: Test whether an input is an Integer and is between two Integers.
; Syntax ........: __LOImpress_IntIsBetween($iTest, $iMin, $iMax[, $vNot = ""[, $vIncl = ""]])
; Parameters ....: $iTest               - an integer value. The Value to test.
;                  $iMin                - an integer value. The minimum $iTest can be.
;                  $iMax                - [optional] an integer value. Default is 0. The maximum $iTest can be.
;                  $vNot                - [optional] a variant value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers inside the min/max range that are not allowed.
;                  $vIncl               - [optional] a variant value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers Outside the min/max range that are allowed.
; Return values .: Success: Boolean
;                  Failure: False and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return Boolean = $iTest not an Integer.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the input is between Min and Max or is an allowed number, and not one of the disallowed numbers, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_IntIsBetween($iTest, $iMin, $iMax = 0, $vNot = "", $vIncl = "")
	If Not IsInt($iTest) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, False)

	Switch @NumParams
		Case 2

			Return SetError($__LO_STATUS_SUCCESS, 0, ($iTest < $iMin) ? (False) : (True))

		Case 3

			Return SetError($__LO_STATUS_SUCCESS, 0, (($iTest < $iMin) Or ($iTest > $iMax)) ? (False) : (True))

		Case 4, 5
			If IsString($vNot) Then
				If StringInStr(":" & $vNot & ":", ":" & $iTest & ":") Then Return SetError($__LO_STATUS_SUCCESS, 0, False)

			ElseIf IsInt($vNot) Then
				If ($iTest = $vNot) Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
			EndIf

			If (($iTest >= $iMin) And ($iTest <= $iMax)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

			If @NumParams = 5 Then ContinueCase

			Return SetError($__LO_STATUS_SUCCESS, 0, False)

		Case Else
			If IsString($vIncl) Then
				If StringInStr(":" & $vIncl & ":", ":" & $iTest & ":") Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

			ElseIf IsInt($vIncl) Then
				If ($iTest = $vIncl) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
			EndIf

			Return SetError($__LO_STATUS_SUCCESS, 0, False)
	EndSwitch
EndFunc   ;==>__LOImpress_IntIsBetween

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_NumIsBetween
; Description ...: Test whether an input is a Number and is between two Numbers.
; Syntax ........: __LOImpress_NumIsBetween($nTest, $nMin, $nMax[, $snNot = ""[, $snIncl = Default]])
; Parameters ....: $nTest               - a general number value. The Value to test.
;                  $nMin                - a general number value. The minimum $iTest can be.
;                  $nMax                - a general number value. The maximum $iTest can be.
;                  $snNot               - [optional] a string value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers inside the min/max range that are not allowed.
;                  $snIncl              - [optional] a string value. Default is Default. Can be a single number, or a String of numbers separated by ":". Defines numbers Outside the min/max range that are allowed.
; Return values .: Success: Boolean
;                  Failure: False
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the input is between Min and Max or is an allowed number, and not one of the disallowed numbers, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_NumIsBetween($nTest, $nMin, $nMax, $snNot = "", $snIncl = Default)
	Local $bMatch = False
	Local $anNot, $anIncl

	If Not IsNumber($nTest) Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
	If (@NumParams = 3) Then Return (($nTest < $nMin) Or ($nTest > $nMax)) ? (SetError($__LO_STATUS_SUCCESS, 0, False)) : (SetError($__LO_STATUS_SUCCESS, 0, True))

	If ($snNot <> "") Then
		If IsString($snNot) And StringInStr($snNot, ":") Then
			$anNot = StringSplit($snNot, ":")
			For $i = 1 To $anNot[0]
				If ($anNot[$i] = $nTest) Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
			Next

		Else
			If ($nTest = $snNot) Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		EndIf
	EndIf

	If (($nTest >= $nMin) And ($nTest <= $nMax)) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	If IsString($snIncl) And StringInStr($snIncl, ":") Then
		$anIncl = StringSplit($snIncl, ":")
		For $j = 1 To $anIncl[0]
			$bMatch = ($anIncl[$j] = $nTest) ? (True) : (False)
			If $bMatch Then ExitLoop
		Next

	ElseIf IsNumber($snIncl) Then
		$bMatch = ($nTest = $snIncl) ? (True) : (False)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $bMatch)
EndFunc   ;==>__LOImpress_NumIsBetween

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_SetPropertyValue
; Description ...: Creates a property value struct object.
; Syntax ........: __LOImpress_SetPropertyValue($sName, $vValue)
; Parameters ....: $sName               - a string value. Property name.
;                  $vValue              - a variant value. Property value.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sName not a string
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Properties Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Property Object Returned
; Author ........: Leagnus, GMK
; Modified ......: donnyh13 - added CreateStruct function. Modified variable names.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_SetPropertyValue($sName, $vValue)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tProperties

	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$tProperties = __LOImpress_CreateStruct("com.sun.star.beans.PropertyValue")
	If @error Or Not IsObj($tProperties) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProperties.Name = $sName
	$tProperties.Value = $vValue

	Return SetError($__LO_STATUS_SUCCESS, 0, $tProperties)
EndFunc   ;==>__LOImpress_SetPropertyValue

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_TransparencyGradientConvert
; Description ...: Convert a Transparency Gradient percentage value to a color value or from a color value to a percentage.
; Syntax ........: __LOImpress_TransparencyGradientConvert([$iPercentToLong = Null[, $iLongToPercent = Null]])
; Parameters ....: $iPercentToLong      - [optional] an integer value. Default is Null. The percentage to convert to Long color integer value.
;                  $iLongToPercent      - [optional] an integer value. Default is Null. The Long color integer value to convert to percentage.
; Return values .: Success: Integer.
;                  Failure: Null and sets the @Error and @Extended flags to non-zero.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return Null = No values called in parameters.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. The requested Integer value converted from percentage to Long color format.
;                  @Error 0 @Extended 1 Return Integer = Success. The requested Integer value from Long color format to percentage.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_TransparencyGradientConvert($iPercentToLong = Null, $iLongToPercent = Null)
	Local $iReturn

	If ($iPercentToLong <> Null) Then
		$iReturn = ((255 * ($iPercentToLong / 100)) + .50) ; Change percentage to decimal and times by White color (255 RGB) Add . 50 to round up if applicable.
		$iReturn = _LOImpress_ConvertColorToLong(Int($iReturn), Int($iReturn), Int($iReturn))

		Return SetError($__LO_STATUS_SUCCESS, 0, $iReturn)

	ElseIf ($iLongToPercent <> Null) Then
		$iReturn = _LOImpress_ConvertColorFromLong(Null, $iLongToPercent)
		$iReturn = Int((($iReturn[0] / 255) * 100) + .50) ; All return color values will be the same, so use only one. Add . 50 to round up if applicable.

		Return SetError($__LO_STATUS_SUCCESS, 1, $iReturn)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, Null)
	EndIf
EndFunc   ;==>__LOImpress_TransparencyGradientConvert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_TransparencyGradientNameInsert
; Description ...: Create and insert a new Transparency Gradient name.
; Syntax ........: __LOImpress_TransparencyGradientNameInsert(ByRef $oDoc, $tTGradient)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $tTGradient          - a dll struct value. A Gradient Structure to copy settings from.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tTGradient not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.TransparencyGradientTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error creating Transparency Gradient Name.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. A new transparency Gradient name was created. Returning the new name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If The Transparency Gradient name is blank, I need to create a new name and apply it. I think I could re-use an old one without problems, but I'm not sure, so to be safe, I will create a new one.
;                  If there are no names that have been already created, then I need to create and apply one before the transparency gradient will be displayed.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_TransparencyGradientNameInsert(ByRef $oDoc, $tTGradient)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tNewTGradient
	Local $oTGradTable
	Local $iCount = 1
	Local $sGradient = "com.sun.star.awt.Gradient2"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tTGradient) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not __LOImpress_VersionCheck(7.6) Then $sGradient = "com.sun.star.awt.Gradient"

	$oTGradTable = $oDoc.createInstance("com.sun.star.drawing.TransparencyGradientTable")
	If Not IsObj($oTGradTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oTGradTable.hasByName("Transparency " & $iCount)
		$iCount += 1
		Sleep((IsInt($iCount / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
	WEnd

	$tNewTGradient = __LOImpress_CreateStruct($sGradient)
	If Not IsObj($tNewTGradient) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; Copy the settings over from the input Style Gradient to my new one. This may not be necessary? But just in case.
	With $tNewTGradient
		.Style = $tTGradient.Style()
		.XOffset = $tTGradient.XOffset()
		.YOffset = $tTGradient.YOffset()
		.Angle = $tTGradient.Angle()
		.Border = $tTGradient.Border()
		.StartColor = $tTGradient.StartColor()
		.EndColor = $tTGradient.EndColor()

		If __LOImpress_VersionCheck(7.6) Then .ColorStops = $tTGradient.ColorStops()
	EndWith

	$oTGradTable.insertByName("Transparency " & $iCount, $tNewTGradient)
	If Not ($oTGradTable.hasByName("Transparency " & $iCount)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, "Transparency " & $iCount)
EndFunc   ;==>__LOImpress_TransparencyGradientNameInsert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_UnitConvert
; Description ...: For converting measurement units.
; Syntax ........: __LOImpress_UnitConvert($nValue, $iReturnType)
; Parameters ....: $nValue              - a general number value. The Number to be converted.
;                  $iReturnType         - a Integer value. Determines conversion type. See Constants, $__LOCONST_CONVERT_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: Integer or Number.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $nValue is not a Number.
;                  @Error 1 @Extended 2 Return 0 = $iReturnType is not a Integer.
;                  @Error 1 @Extended 3 Return 0 = $iReturnType does not match constants, See Constants, $__LOCONST_CONVERT_* as defined in LibreOffice_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 1 Return Number = Returns Number converted from TWIPS to Centimeters.
;                  @Error 0 @Extended 2 Return Number = Returns Number converted from TWIPS to Inches.
;                  @Error 0 @Extended 3 Return Integer = Returns Number converted from Millimeters to uM (Micrometers).
;                  @Error 0 @Extended 4 Return Number = Returns Number converted from Micrometers to MM
;                  @Error 0 @Extended 5 Return Integer = Returns Number converted from Centimeters To uM
;                  @Error 0 @Extended 6 Return Number = Returns Number converted from um (Micrometers) To CM
;                  @Error 0 @Extended 7 Return Integer = Returns Number converted from Inches to uM(Micrometers).
;                  @Error 0 @Extended 8 Return Number = Returns Number converted from uM(Micrometers) to Inches.
;                  @Error 0 @Extended 9 Return Integer = Returns Number converted from TWIPS to uM(Micrometers).
;                  @Error 0 @Extended 10 Return Integer = Returns Number converted from Point to uM(Micrometers).
;                  @Error 0 @Extended 11 Return Number = Returns Number converted from uM(Micrometers) to Point.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_ConvertFromMicrometer, _LOImpress_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_UnitConvert($nValue, $iReturnType)
	Local $iUM, $iMM, $iCM, $iInch

	If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iReturnType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch $iReturnType
		Case $__LOCONST_CONVERT_TWIPS_CM ; TWIPS TO CM
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = ($nValue / 20 / 72)
			; 1 Inch = 2.54 CM
			$iCM = Round(Round($iInch * 2.54, 3), 2)

			Return SetError($__LO_STATUS_SUCCESS, 1, Number($iCM))

		Case $__LOCONST_CONVERT_TWIPS_INCH ; TWIPS to Inch
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = ($nValue / 20 / 72)
			$iInch = Round(Round($iInch, 3), 2)

			Return SetError($__LO_STATUS_SUCCESS, 2, Number($iInch))

		Case $__LOCONST_CONVERT_MM_UM ; Millimeter to Micrometer
			$iUM = ($nValue * 100)
			$iUM = Round(Round($iUM, 1))

			Return SetError($__LO_STATUS_SUCCESS, 3, Number($iUM))

		Case $__LOCONST_CONVERT_UM_MM ; Micrometer to Millimeter
			$iMM = ($nValue / 100)
			$iMM = Round(Round($iMM, 3), 2)

			Return SetError($__LO_STATUS_SUCCESS, 4, Number($iMM))

		Case $__LOCONST_CONVERT_CM_UM ; Centimeter to Micrometer
			$iUM = ($nValue * 1000)
			$iUM = Round(Round($iUM, 1))

			Return SetError($__LO_STATUS_SUCCESS, 5, Int($iUM))

		Case $__LOCONST_CONVERT_UM_CM ; Micrometer to Centimeter
			$iCM = ($nValue / 1000)
			$iCM = Round(Round($iCM, 3), 2)

			Return SetError($__LO_STATUS_SUCCESS, 6, Number($iCM))

		Case $__LOCONST_CONVERT_INCH_UM ; Inch to Micrometer
			; 1 Inch - 2.54 Cm; Micrometer = 1/1000 CM
			$iUM = ($nValue * 2.54) * 1000 ; + .0055
			$iUM = Round(Round($iUM, 1))

			Return SetError($__LO_STATUS_SUCCESS, 7, Int($iUM))

		Case $__LOCONST_CONVERT_UM_INCH ; Micrometer to Inch
			; 1 Inch - 2.54 Cm; Micrometer = 1/1000 CM
			$iInch = ($nValue / 1000) / 2.54 ; + .0055
			$iInch = Round(Round($iInch, 3), 2)

			Return SetError($__LO_STATUS_SUCCESS, 8, $iInch)

		Case $__LOCONST_CONVERT_TWIPS_UM ; TWIPS to Micrometer
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = (($nValue / 20) / 72)
			$iInch = Round(Round($iInch, 3), 2)
			; 1 Inch - 25.4 MM; Micrometer = 1/100 MM
			$iUM = Round($iInch * 25.4 * 100)

			Return SetError($__LO_STATUS_SUCCESS, 9, Int($iUM))

		Case $__LOCONST_CONVERT_PT_UM
			; 1 pt = 35 uM

			Return ($nValue = 0) ? (SetError($__LO_STATUS_SUCCESS, 10, 0)) : (SetError($__LO_STATUS_SUCCESS, 10, Round(($nValue * 35.2778))))

		Case $__LOCONST_CONVERT_UM_PT

			Return ($nValue = 0) ? (SetError($__LO_STATUS_SUCCESS, 11, 0)) : (SetError($__LO_STATUS_SUCCESS, 11, Round(($nValue / 35.2778), 2)))

		Case Else

			Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndSwitch
EndFunc   ;==>__LOImpress_UnitConvert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_VarsAreNull
; Description ...: Tests whether all input parameters are equal to Null keyword.
; Syntax ........: __LOImpress_VarsAreNull($vVar1[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null[, $vVar9 = Null[, $vVar10 = Null[, $vVar11 = Null[, $vVar12 = Null]]]]]]]]]]])
; Parameters ....: $vVar1               - a variant value.
;                  $vVar2               - [optional] a variant value. Default is Null.
;                  $vVar3               - [optional] a variant value. Default is Null.
;                  $vVar4               - [optional] a variant value. Default is Null.
;                  $vVar5               - [optional] a variant value. Default is Null.
;                  $vVar6               - [optional] a variant value. Default is Null.
;                  $vVar7               - [optional] a variant value. Default is Null.
;                  $vVar8               - [optional] a variant value. Default is Null.
;                  $vVar9               - [optional] a variant value. Default is Null.
;                  $vVar10              - [optional] a variant value. Default is Null.
;                  $vVar11              - [optional] a variant value. Default is Null.
;                  $vVar12              - [optional] a variant value. Default is Null.
; Return values .: Success: Boolean
;                  Failure: False
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If All parameters are Equal to Null, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_VarsAreNull($vVar1, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null, $vVar9 = Null, $vVar10 = Null, $vVar11 = Null, $vVar12 = Null)
	Local $bAllNull1, $bAllNull2, $bAllNull3
	$bAllNull1 = (($vVar1 = Null) And ($vVar2 = Null) And ($vVar3 = Null) And ($vVar4 = Null)) ? (True) : (False)
	If (@NumParams <= 4) Then Return SetError($__LO_STATUS_SUCCESS, 0, ($bAllNull1) ? (True) : (False))

	$bAllNull2 = (($vVar5 = Null) And ($vVar6 = Null) And ($vVar7 = Null) And ($vVar8 = Null)) ? (True) : (False)
	If (@NumParams <= 8) Then Return SetError($__LO_STATUS_SUCCESS, 0, ($bAllNull1 And $bAllNull2) ? (True) : (False))

	$bAllNull3 = (($vVar9 = Null) And ($vVar10 = Null) And ($vVar11 = Null) And ($vVar12 = Null)) ? (True) : (False)

	Return SetError($__LO_STATUS_SUCCESS, 0, ($bAllNull1 And $bAllNull2 And $bAllNull3) ? (True) : (False))
EndFunc   ;==>__LOImpress_VarsAreNull

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOImpress_VersionCheck
; Description ...: Test if the currently installed LibreOffice version is high enough to support a certain function.
; Syntax ........: __LOImpress_VersionCheck($fRequiredVersion)
; Parameters ....: $fRequiredVersion    - a floating point value. The version of LibreOffice required.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $fRequiredVersion not a Number.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Current L.O. Version.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the Current L.O. version is higher than or equal to the required version, then True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOImpress_VersionCheck($fRequiredVersion)
	Local Static $sCurrentVersion = _LOImpress_VersionGet(True, False)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, False)

	Local Static $fCurrentVersion = Number($sCurrentVersion)

	If Not IsNumber($fRequiredVersion) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, False)

	Return SetError($__LO_STATUS_SUCCESS, 1, ($fCurrentVersion >= $fRequiredVersion) ? (True) : (False))
EndFunc   ;==>__LOImpress_VersionCheck
