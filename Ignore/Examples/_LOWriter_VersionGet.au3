#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	ObjName_FlagsValue($oServiceManager)

	Local $sVersionAndName, $sFullVersion, $sSimpleVersion

	; Retrieve the current full Office version number and name.
	$sVersionAndName = _LOWriter_VersionGet(False, True)
	If @error Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current full Office version number.
	$sFullVersion = _LOWriter_VersionGet()
	If @error Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current simple Office version number.
	$sSimpleVersion = _LOWriter_VersionGet(True)
	If @error Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Your current full Libre Office version, including the name is: " & $sVersionAndName & @CRLF & _
			"Your current full Libre Office version is: " & $sFullVersion & @CRLF & _
			"Your current simple Libre Office version is: " & $sSimpleVersion)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

Func ObjName_FlagsValue(ByRef $oObj)
        Local $sInfo = ''

        $sInfo &= '+>' & @TAB & 'ObjName($oObj,1) {The name of the Object} =' & @CRLF & @TAB & ObjName($oObj, $OBJ_NAME) & @CRLF

        ; HELPFILE REMARKS: Not all Objects support flags 2 to 7. Always test for @error in these cases.
        $sInfo &= '+>' & @TAB & 'ObjName($oObj,2) {Description string of the Object} =' & @CRLF & @TAB & ObjName($oObj, $OBJ_STRING)
        If @error Then $sInfo &= '@error = ' & @error
        $sInfo &= @CRLF & @CRLF

        $sInfo &= '+>' & @TAB & 'ObjName($oObj,3) {The ProgID of the Object} =' & @CRLF & @TAB & ObjName($oObj, $OBJ_PROGID)
        If @error Then $sInfo &= '@error = ' & @error
        $sInfo &= @CRLF & @CRLF

        $sInfo &= '+>' & @TAB & 'ObjName($oObj,4) {The file that is associated with the object in the Registry} =' & @CRLF & @TAB & ObjName($oObj, $OBJ_FILE)
        If @error Then $sInfo &= '@error = ' & @error
        $sInfo &= @CRLF & @CRLF

        $sInfo &= '+>' & @TAB & 'ObjName($oObj,5) {Module name in which the object runs (WIN XP And above). Marshaller for non-inproc objects.} =' & @CRLF & @TAB & ObjName($oObj, $OBJ_MODULE)
        If @error Then $sInfo &= '@error = ' & @error
        $sInfo &= @CRLF & @CRLF

        $sInfo &= '+>' & @TAB & 'ObjName($oObj,6) {CLSID of the object''s coclass} =' & @CRLF & @TAB & ObjName($oObj, $OBJ_CLSID)
        If @error Then $sInfo &= '@error = ' & @error
        $sInfo &= @CRLF & @CRLF

        $sInfo &= '+>' & @TAB & 'ObjName($oObj,7) {IID of the object''s interface} =' & @CRLF & @TAB & ObjName($oObj, $OBJ_IID)
        If @error Then $sInfo &= '@error = ' & @error
        $sInfo &= @CRLF & @CRLF

        MsgBox($MB_SYSTEMMODAL, "ObjName:", $sInfo)
EndFunc   ;==>ObjName_FlagsValue
