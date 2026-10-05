#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#SciTE4AutoIt3_Dynamic_Include_Path="C:\Users\Dave\Documents\GitHub\donnyh13\Au3LibreOffice"
#Tidy_Parameters=/sefc /reel /bdir="C:\Users\Dave\Documents\GitHub\Tidy_Backups\"

#Region ; *** Dynamically added Include files ***
#include "C:\Users\Dave\Documents\GitHub\donnyh13\Au3LibreOffice\LibreOfficeImpress_Constants.au3" ; added:12/27/23 06:54:24
#include "C:\Users\Dave\Documents\GitHub\donnyh13\Au3LibreOffice\LibreOfficeImpress_Cursor.au3" ; added:12/27/23 06:54:24
#include "C:\Users\Dave\Documents\GitHub\donnyh13\Au3LibreOffice\LibreOfficeImpress_Helper.au3" ; added:12/31/23 12:23:37
#include "C:\Users\Dave\Documents\GitHub\donnyh13\Au3LibreOffice\LibreOfficeImpress_Doc.au3" ; added:12/31/23 12:23:37
#include "C:\Users\Dave\Documents\GitHub\donnyh13\Au3LibreOffice\LibreOfficeImpress_Field.au3" ; added:12/31/23 12:23:37
#include "C:\Users\Dave\Documents\GitHub\donnyh13\Au3LibreOffice\LibreOfficeImpress_Shape.au3" ; added:12/31/23 12:23:37
#include "C:\Users\Dave\Documents\GitHub\donnyh13\Au3LibreOffice\LibreOfficeImpress_Slide.au3" ; added:12/31/23 12:23:37
#include "C:\Users\Dave\Documents\GitHub\donnyh13\Au3LibreOffice\LibreOfficeImpress_Internal.au3" ; added:12/31/23 12:23:37
#EndRegion ; *** Dynamically added Include files ***
_LOImpress_ComError_UserFunction(ConsoleWrite)

#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape, $oTextBox, $oGroup
	Local $aoShapes[2]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active Slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_BLANK
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_BLANK)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Textbox.
	$oTextBox = _LOImpress_ShapeTextBoxInsert($oSlide, $LOI_SHAPE_TEXTBOX_TYPE_TEXTBOX, 10500, 5000, -1, 1500)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a line around the Text Box so it's visible.
	_LOImpress_ShapeLineProperties($oTextBox, $LOI_SHAPE_LINE_STYLE_CONTINUOUS)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Page, 3000 Wide by 6000 High.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000, 2000, 3500)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add both shapes to an array.
	$aoShapes[0] = $oTextBox
	$aoShapes[1] = $oShape

	; Create a group containing both the Textbox and rectangle.
	$oGroup = _LOImpress_ShapeGroupCreate($oSlide, $aoShapes)
	If @error Then _ERROR($oDoc, "Failed to create a shape group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have created a group. Press ok to ungroup the group.")

	; Ungroup the group of shapes.
	_LOImpress_ShapeGroupDelete($oGroup)
	If @error Then _ERROR($oDoc, "Failed to ungroup a group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
