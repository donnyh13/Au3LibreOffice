#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#Tidy_Parameters=/reel /tcl=1

#include <ColorConstants.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <StringConstants.au3>
#include <WindowsNotifsConstants.au3>
#include <WindowsStylesConstants.au3>

#Region Header

; #INDEX# =======================================================================================================================
; Title .........: Simple Library Docs Generator
; AutoIt Version : v3.3+
; Description ...: Create and compile help files for AutoIt functions.
; Author(s) .....: G.Sandler (a.k.a (Mr)CreatoR)
; Modified ......: water
; Modified ......: donnyh13 -- Added multi-file support, customized file hierarchy and styling for Au3LibreOffice, added a few more options, integrated Scite syntax highlighting by Jos and guinness, etc.
; Dll ...........:
; Link ..........: https://www.autoitscript.com/forum/topic/207211-ghs-thread-for-the-modified-simple-library-docs-generator/
;
; ===============================================================================================================================

#EndRegion Header

#NoTrayIcon

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#Region Global Variables

Global $sSciTECmd
Global $hSciTE_hwnd = WinGetHandle("DirectorExtension")
Global $sCurrentFile

Global $sAppName = "Simple Library Docs Generator"
Global $sAppSettingFile = @ScriptDir & "\" & $sAppName & " Settings.ini"
If Not FileExists($sAppSettingFile) Then FileWrite($sAppSettingFile, "") ; Create the setting file if it doesn't exist.

;Читаем последние введённые данные программы ; Read the last entered program data
Global $sLibraryFile = IniRead($sAppSettingFile, "Settings", "Library File", "")
Global $sLibExamplesDir = IniRead($sAppSettingFile, "Settings", "Library Examples Dir", StringRegExpReplace($sLibraryFile, "\\[^\\]*$", ""))
Global $sDocsDir = IniRead($sAppSettingFile, "Settings", "Docs Dir", "")
Global $sDocsIndex_Title = IniRead($sAppSettingFile, "Settings", "Docs Index Title", "UDF Library")
Global $sDocsIndexHeader_Title = IniRead($sAppSettingFile, "Settings", "Docs Index Header Title", "UDF Library")
Global $sDocsFunc_Title = IniRead($sAppSettingFile, "Settings", "Docs Func Title", "Function Reference [%IndexTitle% - %FuncName%]")
Global $sDocsFuncHeader_Title = IniRead($sAppSettingFile, "Settings", "Docs Func Header Title", "Function Reference")
Global $sCHMFileName = IniRead($sAppSettingFile, "Settings", "CHM File Name", $sDocsIndex_Title)
Global $iShowGeneratedDocs = IniRead($sAppSettingFile, "Settings", "Show Generated Docs", $GUI_UNCHECKED)
Global $iHighlightExampleSyntax = IniRead($sAppSettingFile, "Settings", "Highlight Example Syntax", $GUI_UNCHECKED)
Global $iHiglExmplSyntx_AddURLs = IniRead($sAppSettingFile, "Settings", "Highlight Example Syntax Add URLs", $GUI_CHECKED)
Global $iCompileToChm = IniRead($sAppSettingFile, "Settings", "Compile To Chm", $GUI_UNCHECKED)
Global $iCompile_DelOnDone = IniRead($sAppSettingFile, "Settings", "Delete Source on Compile", $GUI_UNCHECKED)

Global $hGUI
Global $asSplit
Global $idLibraryFile_Button, $idLibraryFile_Input, $idDocsDir_Button, $idDocsDir_Input, $idLibExamplesDir_Button, $idLibExamplesDir_Input, $idDocsIndexTitle_Input, _
		$idDocsIndexHeaderTitle_Input, $idDocsFuncTitle_Input, $idDocsFuncHeaderTitle_Input, $idShowGeneratedDocs_CB, $idHighlightExampleSyntax_CB, _
		$idHiglExmplSyntx_AddURLs_CB, $idCompileToChm_CB, $idCompile_DelOnDone_CB, $idStatus_Label, $idGenerateDocs_Button, $idExit_Button, $idCHMFileName_Input
Global $iTempCount = 1
Global $sInitial_Dir, $sDir, $sLibraryFile_Name, $sFile, $sIndex_File, $sLibraryDir, $sTempFolder
Global Enum $__g_Header_Name, $__g_Header_Description, $__g_Header_Syntax, $__g_Header_Parameters, $__g_Header_Return, $__g_Header_Author, $__g_Header_Modified, _
		$__g_Header_Remarks, $__g_Header_Related, $__g_Header_Link, $__g_Header_Example

#EndRegion Global Variables

#Region GUI Creation & Main Loop
$hGUI = GUICreate($sAppName, 400, 510, -1, -1, -1, $WS_EX_TOPMOST)

GUICtrlCreateLabel("Au3 Library File(s):", 20, 5, -1, 15)
$idLibraryFile_Button = GUICtrlCreateButton("...", 360, 20, 20, 20)
$idLibraryFile_Input = GUICtrlCreateInput($sLibraryFile, 20, 20, 340, 20, $ES_READONLY)

GUICtrlCreateLabel("Au3 Library Examples Dir:", 20, 45, -1, 15)
$idLibExamplesDir_Button = GUICtrlCreateButton("...", 360, 60, 20, 20)
$idLibExamplesDir_Input = GUICtrlCreateInput($sLibExamplesDir, 20, 60, 340, 20, $ES_READONLY)

GUICtrlCreateLabel("Au3 Library Documents Output Dir:", 20, 85, -1, 15)
$idDocsDir_Button = GUICtrlCreateButton("...", 360, 100, 20, 20)
$idDocsDir_Input = GUICtrlCreateInput($sDocsDir, 20, 100, 340, 20)

GUICtrlCreateGroup("Title Formats", 10, 123, 380, 220)

GUICtrlCreateLabel("Document Index Title" & @CRLF & "(Browser Title)", 20, 140, -1, 27)
$idDocsIndexTitle_Input = GUICtrlCreateInput($sDocsIndex_Title, 20, 170, 170, 20)

GUICtrlCreateLabel("Document Index Header Title" & @CRLF & "(<h1> element)", 210, 140, -1, 27)
$idDocsIndexHeaderTitle_Input = GUICtrlCreateInput($sDocsIndexHeader_Title, 210, 170, 170, 20)

GUICtrlCreateLabel("Document Function Title" & @CRLF & "(Browser Title)", 20, 200, -1, 27)
$idDocsFuncTitle_Input = GUICtrlCreateInput($sDocsFunc_Title, 20, 230, 170, 20)

GUICtrlCreateLabel("Document Function Header Title" & @CRLF & "(<h1> element)", 210, 200, -1, 27)
$idDocsFuncHeaderTitle_Input = GUICtrlCreateInput($sDocsFuncHeader_Title, 210, 230, 170, 20)

GUICtrlCreateLabel( _
		"%IndexTitle% 		= Document Index Title" & @CRLF & _
		"%IndexHeaderTitle% 	= Document Index Header Title" & @CRLF & _
		"%FuncName% 		= Function Name (current function in the list)", 20, 255, 360, 37)

GUICtrlSetFont(-1, 8, 200, 0, "Tahoma")
GUICtrlSetColor(-1, $COLOR_BLUE)

GUICtrlCreateLabel("CHM Output Name", 20, 300, 100, 15)
$idCHMFileName_Input = GUICtrlCreateInput($sCHMFileName, 20, 315, 170, 20)

GUICtrlCreateGroup("Options", 10, 350, 380, 125)

$idShowGeneratedDocs_CB = GUICtrlCreateCheckbox("Show Generated Docs", 20, 370, -1, 15)
GUICtrlSetState($idShowGeneratedDocs_CB, $iShowGeneratedDocs)

$idHighlightExampleSyntax_CB = GUICtrlCreateCheckbox("Highlight Example Syntax (!Not Recomended -> longer than usual)", 20, 390, -1, 15)
GUICtrlSetState($idHighlightExampleSyntax_CB, $iHighlightExampleSyntax)

$idHiglExmplSyntx_AddURLs_CB = GUICtrlCreateCheckbox("Add URLs to online docs for built-in functions", 40, 410, -1, 15)
GUICtrlSetState($idHiglExmplSyntx_AddURLs_CB, $iHiglExmplSyntx_AddURLs)
If (BitAND(GUICtrlRead($idHighlightExampleSyntax_CB), $GUI_UNCHECKED) = $GUI_UNCHECKED) Then GUICtrlSetState($idHiglExmplSyntx_AddURLs_CB, $GUI_DISABLE)

$idCompileToChm_CB = GUICtrlCreateCheckbox("Compile Document To Chm Help File", 20, 430, -1, 15)
GUICtrlSetState($idCompileToChm_CB, $iCompileToChm)

$idCompile_DelOnDone_CB = GUICtrlCreateCheckbox("Delete Build Files On Finish", 40, 450, -1, 15)
GUICtrlSetState($idCompile_DelOnDone_CB, $iCompile_DelOnDone)
If (BitAND(GUICtrlRead($idCompileToChm_CB), $GUI_UNCHECKED) = $GUI_UNCHECKED) Then GUICtrlSetState($idCompile_DelOnDone_CB, $GUI_DISABLE)

$idStatus_Label = GUICtrlCreateLabel("", 230, 480, 160, 27)
GUICtrlSetFont(-1, 8.5, 800, 0, "Georgia")
GUICtrlSetColor(-1, $COLOR_RED)

$idGenerateDocs_Button = GUICtrlCreateButton("Generate", 20, 480, 60, 20)
$idExit_Button = GUICtrlCreateButton("Exit", 100, 480, 60, 20)

GUISetState(@SW_SHOW, $hGUI)

While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE, $idExit_Button
			_SaveSettings()
			Exit
		Case $idLibraryFile_Button
			$sInitial_Dir = GUICtrlRead($idLibraryFile_Input)
			If StringInStr($sInitial_Dir, "|") Then $sInitial_Dir = StringLeft($sInitial_Dir, StringInStr($sInitial_Dir, "|") - 1)
			If Not FileExists($sInitial_Dir) Then $sInitial_Dir = @WorkingDir
			$sFile = FileOpenDialog("Select Au3 Library File", $sInitial_Dir, "AutoIt Script File (*.au3)", BitOR($FD_FILEMUSTEXIST, $FD_MULTISELECT), "", $hGUI)
			If @error Then ContinueLoop

			GUICtrlSetData($idLibraryFile_Input, $sFile)
		Case $idLibExamplesDir_Button
			$sInitial_Dir = GUICtrlRead($idLibExamplesDir_Input)
			If Not FileExists($sInitial_Dir) Then $sInitial_Dir = @WorkingDir

			$sDir = FileSelectFolder("Select Au3 Library Examples Directory", "", $FSF_CREATEBUTTON + $FSF_NEWDIALOG, $sInitial_Dir, $hGUI)
			If @error Or Not StringInStr($sDir, "\") Or Not FileExists($sDir) Then ContinueLoop

			GUICtrlSetData($idLibExamplesDir_Input, $sDir)
		Case $idDocsDir_Button
			$sInitial_Dir = GUICtrlRead($idDocsDir_Input)
			If Not FileExists($sInitial_Dir) Then $sInitial_Dir = @WorkingDir

			$sDir = FileSelectFolder("Select Au3 Documents Output Directory", "", $FSF_CREATEBUTTON + $FSF_NEWDIALOG, $sInitial_Dir, $hGUI)
			If @error Or Not StringInStr($sDir, "\") Or Not FileExists($sDir) Then ContinueLoop

			GUICtrlSetData($idDocsDir_Input, $sDir)
		Case $idHighlightExampleSyntax_CB
			If GUICtrlRead($idHighlightExampleSyntax_CB) = $GUI_CHECKED Then
				GUICtrlSetState($idHiglExmplSyntx_AddURLs_CB, $GUI_ENABLE)
			Else
				GUICtrlSetState($idHiglExmplSyntx_AddURLs_CB, $GUI_DISABLE)
				If (GUICtrlRead($idHiglExmplSyntx_AddURLs_CB) = $GUI_CHECKED) Then GUICtrlSetState($idHiglExmplSyntx_AddURLs_CB, $GUI_UNCHECKED)
			EndIf
		Case $idCompileToChm_CB
			If GUICtrlRead($idCompileToChm_CB) = $GUI_CHECKED Then
				GUICtrlSetState($idCompile_DelOnDone_CB, $GUI_ENABLE)
			Else
				GUICtrlSetState($idCompile_DelOnDone_CB, $GUI_DISABLE)
				If (GUICtrlRead($idCompile_DelOnDone_CB) = $GUI_CHECKED) Then GUICtrlSetState($idCompile_DelOnDone_CB, $GUI_UNCHECKED)
			EndIf
		Case $idGenerateDocs_Button
			$sLibraryFile = GUICtrlRead($idLibraryFile_Input)
			$sLibExamplesDir = GUICtrlRead($idLibExamplesDir_Input)

			$sDocsDir = GUICtrlRead($idDocsDir_Input)
			$sDocsIndex_Title = GUICtrlRead($idDocsIndexTitle_Input)
			$sDocsIndexHeader_Title = GUICtrlRead($idDocsIndexHeaderTitle_Input)
			$sDocsFunc_Title = GUICtrlRead($idDocsFuncTitle_Input)
			$sDocsFuncHeader_Title = GUICtrlRead($idDocsFuncHeaderTitle_Input)

			$iShowGeneratedDocs = Number(GUICtrlRead($idShowGeneratedDocs_CB) = $GUI_CHECKED)
			$iHighlightExampleSyntax = Number(GUICtrlRead($idHighlightExampleSyntax_CB) = $GUI_CHECKED)
			$iHiglExmplSyntx_AddURLs = Number(GUICtrlRead($idHiglExmplSyntx_AddURLs_CB) = $GUI_CHECKED)
			$iCompileToChm = Number(GUICtrlRead($idCompileToChm_CB) = $GUI_CHECKED)
			$iCompile_DelOnDone = Number(GUICtrlRead($idCompile_DelOnDone_CB) = $GUI_CHECKED)

			; Register COPYDATA message if we need to use Scite to style the examples.
			If $iHighlightExampleSyntax And IsHWnd($hSciTE_hwnd) Then
				GUIRegisterMsg($WM_COPYDATA, "_MY_WM_COPYDATA")
				$sCurrentFile = StringReplace(_SendSciTE_GetInfo($hGUI, $hSciTE_hwnd, "askfilename:"), "\\", "\")
				$sCurrentFile = StringReplace($sCurrentFile, "filename:", "")
			EndIf

			If $iCompile_DelOnDone Then ; Create a temporary folder to output files to.
				$sTempFolder = "\Au3LibOut-" & @MDAY & "-" & @MIN & "-"
				$iTempCount = 1
				While FileExists($sDocsDir & $sTempFolder & $iTempCount)
					$iTempCount += 1
					Sleep(10)
				WEnd
				$sTempFolder = $sTempFolder & $iTempCount
				$sDocsDir &= $sTempFolder ; Add the temp folder to the path.
				DirCreate($sDocsDir)
			EndIf

			$asSplit = StringSplit($sLibraryFile, "|", $STR_NOCOUNT)

			$sLibraryDir = "" ; Clear any former library directory.

			For $iFile = 0 To UBound($asSplit) - 1
				If ($iFile > 0) Then ; If Multiple files are selected, add the path (in element 0) to each name.
					$asSplit[$iFile] = $asSplit[0] & "\" & $asSplit[$iFile]

				EndIf

				If Not FileExists($asSplit[$iFile]) Then
					MsgBox($MB_ICONWARNING, 'Attention', 'Library File Does Not Exist: ' & $asSplit[$iFile], 0, $hGUI)
					ContinueLoop 2
				EndIf
			Next

			If (UBound($asSplit) > 1) Then
				$sLibraryDir = $asSplit[0] ; Save the directory path.
				$asSplit[0] = "" ; Clear the path when there are multiple files, as it has been added to each single file and saved to LibraryDir.

			Else
				$sLibraryDir = StringLeft($asSplit[0], StringInStr($asSplit[0], "\", 0, -1) - 1) ; There's only one path so grab the directory path from the full path.
			EndIf

			$sLibraryFile_Name = GUICtrlRead($idCHMFileName_Input)
			If ($sLibraryFile_Name = "") Then $sLibraryFile_Name = "OutPut"

			GUICtrlSetData($idStatus_Label, "Please wait... " & @CRLF & "[Generating Docs]")
			GUICtrlSetState($idGenerateDocs_Button, $GUI_DISABLE)
			GUICtrlSetState($idExit_Button, $GUI_DISABLE)

			;Включаем режим событий (OnEvent) чтобы корректно обрабатывать пункт "Highlight Example Syntax" и "Add URLs to online docs"
			; Enable OnEvent mode so "Highlight Example Syntax" and "Add URLs to online docs" are handled correctly.

;~ 			Opt("GUIOnEventMode", 1)

			;Задаём событие для этих пунктов
			; Assign event handlers for these controls.
;~ 			GUICtrlSetOnEvent($idHighlightExampleSyntax_CB, "_HighlightExampleSyntax_CB_Event")
;~ 			GUICtrlSetOnEvent($idHiglExmplSyntx_AddURLs_CB, "_HighlightExampleSyntax_CB_Event")

			$sIndex_File = _AU3Lib_GenerateDocs_Proc($asSplit)

			;Отключаем режим событий (OnEvent)
			; Disable OnEvent mode.
;~ 			Opt("GUIOnEventMode", 0)

			If $iCompileToChm Then
				$sIndex_File = $sDocsDir & "\" & $sLibraryFile_Name & ".chm"
				_AU3Lib_CompileHtmlToChm($sDocsDir, $sLibraryFile_Name)
				; Delete the temporary folder if we made one, and only if moving the CHM was successful.
				If $iCompile_DelOnDone And FileMove($sDocsDir & "\" & $sLibraryFile_Name & ".chm", StringReplace($sDocsDir, $sTempFolder, "") & "\" & $sLibraryFile_Name & ".chm", $FC_OVERWRITE) Then DirRemove($sDocsDir, $DIR_REMOVE)

				; Remove the temp folder from the index path if the CHM move was successful, in case Show generated docs is checked.
				If $iCompile_DelOnDone And FileExists(StringReplace($sDocsDir, $sTempFolder, "") & "\" & $sLibraryFile_Name & ".chm") Then $sIndex_File = StringReplace($sIndex_File, $sTempFolder, "")
			EndIf

			If $iShowGeneratedDocs Then
				ShellExecute($sIndex_File)
			EndIf

			GUICtrlSetData($idStatus_Label, @CRLF & "Done!")
			GUICtrlSetState($idGenerateDocs_Button, $GUI_ENABLE)
			GUICtrlSetState($idExit_Button, $GUI_ENABLE)

			; UnRegister COPYDATA message.
			If $iHighlightExampleSyntax Then GUIRegisterMsg($WM_COPYDATA, "")

			;Сохранение данных
			; Save settings/data.
			_SaveSettings()
	EndSwitch
WEnd
#EndRegion GUI Creation & Main Loop

#Region Core Functions
Func _AU3Lib_GenerateDocs_Proc(ByRef $asFiles)
	;Объявляем переменные
	; Declare local variables.
	Local $aHeaders, $sIndex_File, $hIndex_File, $sHtml_Header, $sFuncName, $aParams_Split, $aReturn, $sReturn_Success, $sReturn_Failure, $sRemarks, $aSplit, $sRelated
	Local $sHtml_Content, $sParam_TD_Width, $sValue_TD_Width, $sExample_Content, $hFunc_File, $sFile
	Local $sDocsFunc_Formatted_Title, $sDocsFuncHeader_Formatted_Title, $sDocsFunc_FormattedTmp_Title, $sDocsFuncHeader_FormattedTmp_Title ;, $sCurrWorkDir

	Local $asRegExp[0], $asExamples[0]
	Local $aParams[0], $aValues[0]
	Local $hCompIndex, $hSubCompIndex, $hSearch, $hTempFile
	Local $avIndexFiles[0][3], $avSubIndex[0]
	Local $iCompIndex = -1
	Local $sComponent = "", $sSubComponent = "", $sCompHTML_Header = "", $sSubCompHTML_Header = "", $sDesc, $sTempExamplePath, $sTempExampleName, $sFileName
	Local Enum $__eIndexName, $__eIndexHandle, $__eIndexSubIndex
	Local $sTempFile = _TempFile(@TempDir, "1", ".au3"), $sTempOutputFile = _TempFile(@TempDir, "1", ".html")

	;Форматируем заголовки для Html-страниц
	; Format the HTML page titles.
	$sDocsFunc_Formatted_Title = StringReplace($sDocsFunc_Title, '%IndexTitle%', $sDocsIndex_Title)
	$sDocsFunc_Formatted_Title = StringReplace($sDocsFunc_Formatted_Title, '%IndexHeaderTitle%', $sDocsIndexHeader_Title)
	$sDocsFuncHeader_Formatted_Title = StringReplace($sDocsFuncHeader_Title, '%IndexTitle%', $sDocsIndex_Title)
	$sDocsFuncHeader_Formatted_Title = StringReplace($sDocsFuncHeader_Formatted_Title, '%IndexHeaderTitle%', $sDocsIndexHeader_Title)

	;Создаём структуру каталогов
	; Create the output directory structure.
	DirCreate($sDocsDir & "\css")
	DirCreate($sDocsDir & "\funcs")
	DirCreate($sDocsDir & "\indices")

;~ 	$sCurrWorkDir = @WorkingDir ; Backup the working Directory.

	; Change the working Directory to the ScriptDir so that the files can be found.
	If (@WorkingDir <> @ScriptDir) Then FileChangeDir(@ScriptDir)

	;Добавляем файл стилей и картинку для <h1>
	; Copy the stylesheet and background image for the <h1> header.
	FileInstall(".\Resources\default.css", $sDocsDir & "\css\default.css", $FC_OVERWRITE)
	FileInstall(".\Resources\h1_background.jpg", $sDocsDir & "\css\h1_background.jpg", $FC_OVERWRITE)

	; Change the Working Directory back to the former value.
;~ 	FileChangeDir($sCurrWorkDir)

	;Определяем файл индекса
	; Create the index file.
	$sIndex_File = $sDocsDir & "\index.htm"

	;Открываем файл индекса (в UTF-8)
	; Open (and create) the index file (UTF-8).
	$hIndex_File = FileOpen($sIndex_File, $FO_OVERWRITE + $FO_UTF8)

	;Получаем главный заголовок Html-файла
	; Generate the HTML page header.
	$sHtml_Header = _AU3Lib_GetHtmlHeaderStr($sDocsIndex_Title, 'css')

	;Пишем сразу заголовок индекса
	; Write the beginning of the index page.
	FileWriteLine($hIndex_File, $sHtml_Header & @CRLF & '<body>' & @CRLF & '  <h1>' & $sDocsIndexHeader_Title & '</h1>' & @CRLF & _
			'  <p>Below is a complete list of the LibreOffice Components available in this library. Click on a Component name for a list of available Sub-Components.</p>')

	For $iFile = 0 To UBound($asFiles) - 1
		; Skip and Null files that are not valid, Internal, Constants, and top level files.
		; e.g.: LibreOffice.au3, LibreOffice_Constants.au3, LibreOffice_Internal.au3
;~ 		If Not StringInStr($asFiles[$iFile], "_") Or StringRegExp($asFiles[$iFile], "(?i)_(?:Constants|Internal)") Then
		If Not StringRegExp($asFiles[$iFile], "[^\\]+_[^\\]+$") Or StringRegExp($asFiles[$iFile], "(?i)[^\\]+_(?:Constants|Internal)[^\\]+$") Then
			$asFiles[$iFile] = "" ; Scrap the file name so that the array is only filled with valid data.
		EndIf
	Next

	; Parse and store all function names to help validate links.
	__LinkIsValid(Default, $asFiles)

	Local $iTotalFiles = UBound($asFiles)
	If ($iTotalFiles > 1) Then $iTotalFiles -= 1 ; Minus one from the count since the file path was part of the array.

	For $iFile = 0 To UBound($asFiles) - 1
		GUICtrlSetData($idStatus_Label, "Please wait..." & @CRLF & "[Generating Docs " & $iFile & " of " & $iTotalFiles & "]")
		If ($asFiles[$iFile] = "") Then ContinueLoop

		$iCompIndex = -1

		; Determine the Index file to use, or create it if it doesn't exist.
;~ 		$asRegExp = StringRegExp($asFiles[$iFile], "(?i)(?|LibreOffice(\w+)_(\w+)|(LibreOffice)_(\w+))", $STR_REGEXPARRAYMATCH) ; Get Sub-Component name, else get LibreOffice name for global files.
		$asRegExp = StringRegExp($asFiles[$iFile], "(?i)(?|LibreOffice(\w+)_(\w+)|(LibreOffice)_(\w+))[^\\]+$", $STR_REGEXPARRAYMATCH) ; Get Sub-Component name, else get LibreOffice name for global files.
		If Not IsArray($asRegExp) Then
			ConsoleWrite("! Failed to identify Component and Sub-Component, skipping: " & $asFiles[$iFile] & @CRLF)
			ContinueLoop
		EndIf

		$sComponent = ($asRegExp[0] = "LibreOffice") ? ("Global") : ($asRegExp[0]) ; If dealing with the global helper funcs (LO_Helper) etc., just name it global.
		$sSubComponent = $asRegExp[1]

		For $i = 0 To UBound($avIndexFiles) - 1
			If ($avIndexFiles[$i][$__eIndexName] = $sComponent & "_Index") Then
				$iCompIndex = $i
				$avSubIndex = $avIndexFiles[$i][$__eIndexSubIndex]
				ExitLoop
			EndIf
		Next

		; Create the main component index file if not found.
		If ($iCompIndex = -1) Then
			; Add the component to the main index.
			FileWriteLine($hIndex_File, '  <ul>' & @CRLF & _
					'    <li><a href="indices/' & $sComponent & '_Index' & '.htm">' & $sComponent & '</a></li>' & @CRLF & _
					'  </ul>')

			; Open the index file (UTF-8).
			$hCompIndex = FileOpen($sDocsDir & "\indices\" & $sComponent & "_Index.htm", $FO_OVERWRITE + $FO_UTF8)

			; Generate the HTML page header.
			$sCompHTML_Header = _AU3Lib_GetHtmlHeaderStr($sDocsIndex_Title & " " & $sComponent, '../css')

			; Write the beginning of the Component index page.
			FileWriteLine($hCompIndex, $sCompHTML_Header & @CRLF & '<body>' & @CRLF & '  <h1>' & "LibreOffice " & $sComponent & '</h1>' & @CRLF & _
					'  <p>Below is a complete list of the sub-modules available for the LibreOffice ' & $sComponent & ' module. Click on a module name for a complete list of functions.</p>' & @CRLF & _
					'  <table>' & @CRLF & _
					'  <tr>' & @CRLF & _
					'    <th style="width:25%">Sub-Module</th>' & @CRLF & _
					'    <th style="width:75%">Description</th>' & @CRLF & _
					'  </tr>')

			$iCompIndex = UBound($avIndexFiles)

			ReDim $avIndexFiles[$iCompIndex + 1][3]
			ReDim $avSubIndex[0]

			$avIndexFiles[$iCompIndex][$__eIndexName] = $sComponent & "_Index"
			$avIndexFiles[$iCompIndex][$__eIndexHandle] = $hCompIndex
			$avIndexFiles[$iCompIndex][$__eIndexSubIndex] = $avSubIndex
			$avSubIndex = $avIndexFiles[$iCompIndex][$__eIndexSubIndex]
		EndIf

		; Create the Sub-Component Index
		ReDim $avSubIndex[UBound($avSubIndex) + 1]
		$avSubIndex[UBound($avSubIndex) - 1] = $sComponent & "_" & $sSubComponent & "_Index"
		; Open the index file (UTF-8).
		$hSubCompIndex = FileOpen($sDocsDir & "\indices\" & $avSubIndex[UBound($avSubIndex) - 1] & ".htm", $FO_OVERWRITE + $FO_UTF8)

		; Generate the HTML page header.
		$sSubCompHTML_Header = _AU3Lib_GetHtmlHeaderStr($sComponent & " " & $sSubComponent, '../css')

		; Write the beginning of the index page.
		FileWriteLine($hSubCompIndex, $sSubCompHTML_Header & @CRLF & '<body>' & @CRLF & '  <h1>' & 'LibreOffice ' & $sComponent & ' ' & $sSubComponent & '</h1>' & @CRLF & _
				'  <p>Below is a complete list of the functions available in the LibreOffice ' & $sComponent & ' ' & $sSubComponent & ' library. Click on a function name for a detailed description.</p>' & @CRLF & _
				'  <table>' & @CRLF & _
				'  <tr>' & @CRLF & _
				'    <th style="width:25%">Function</th>' & @CRLF & _
				'    <th style="width:75%">Description</th>' & @CRLF & _
				'  </tr>')

		; Grab the Description for the Sub-Component from the file.
		$asRegExp = StringRegExp(FileRead($asFiles[$iFile]), "(?im)^;\s*Description\s*\.+:\s*(.+)$", $STR_REGEXPARRAYMATCH)

		; If no description is found, use a blank string
		If Not IsArray($asRegExp) Then
			$sDesc = ""
		Else
			$sDesc = $asRegExp[0]
		EndIf

		; Add the Sub-Component index entry to the Component Index.
		FileWriteLine($hCompIndex, '  <tr>' & @CRLF & _
				'    <td><a href="' & $avSubIndex[UBound($avSubIndex) - 1] & '.htm">' & $sSubComponent & '</a></td>' & @CRLF & _
				'    <td>' & $sDesc & '</td>' & @CRLF & _ ; Add the description.
				'  </tr>')

		;Получаем заголовки библиотеки
		; Retrieve the library headers.
		$aHeaders = _AU3Lib_GetHeaders($asFiles[$iFile])

		For $i = 1 To UBound($aHeaders) - 1
			;Получаем имя функций (для имени файла и ссылки)
			; Get the function name (used for the filename and hyperlink).
			$sFuncName = $aHeaders[$i][$__g_Header_Name]

			;Форматируем заголовки для Html-страниц (только для функций, т.к. она изменяется и считывается в цикле)
			; Format the HTML titles for this function (updated for each function in the loop).
			$sDocsFunc_FormattedTmp_Title = StringReplace($sDocsFunc_Formatted_Title, '%FuncName%', $sFuncName)
			$sDocsFuncHeader_FormattedTmp_Title = StringReplace($sDocsFuncHeader_Formatted_Title, '%FuncName%', $sFuncName)

			;Получаем имя файла функций
			; Generate the function filename.
			$sFile = _AU3Lib_GetValidFileName($sFuncName)

			; Add the current function entry to the Sub-Component index.
			FileWriteLine($hSubCompIndex, '  <tr>' & @CRLF & _
					'    <td><a href="../funcs/' & $sFile & '.htm">' & $sFuncName & '</a></td>' & @CRLF & _
					'    <td>' & $aHeaders[$i][$__g_Header_Description] & '</td>' & @CRLF & _
					'  </tr>')

			#Region Получаем и форматируем параметры функций ; Get and format function parameters
			$aParams_Split = StringSplit(StringStripCR($aHeaders[$i][$__g_Header_Parameters]), @LF)
;~ 			Local $aParams[$aParams_Split[0] + 1]
			ReDim $aParams[$aParams_Split[0] + 1]
			$aParams[0] = 0 ; Reset the value to 0 so the correct count is set.
;~ 			Local $aValues[$aParams_Split[0] + 1]
			ReDim $aValues[$aParams_Split[0] + 1]
			$aValues[0] = 0 ; Reset the value to 0 so the correct count is set.

			For $j = 1 To $aParams_Split[0]
				If StringIsSpace($aParams_Split[$j]) Then ContinueLoop

				;Проверяем если текущая строка это строка с параметром в начале
				; Check whether the current line starts with a parameter definition.
				If StringRegExp($aParams_Split[$j], "(?:\s+)?(\$[\w\d_]+)\s+-.*") Then
					$aParams[0] += 1
					$aValues[0] += 1

					;Извлекаем параметр
					; Extract the parameter name.
					$aParams[$aParams[0]] = StringRegExpReplace($aParams_Split[$j], '(?:\s+)?(\$[\w\d_]+)\s+-.*', '\1')

					;Извлекаем описание параметра
					; Extract the parameter description.
					$aValues[$aValues[0]] = StringRegExpReplace($aParams_Split[$j], '(?:\s+)?\$[\w\d_]+\s+-\s+(.*)', '\1') & '<br>' & @CRLF
				Else ;Иначе это продолжение описания параметра
					; Otherwise this line continues the parameter description.
					$aValues[$aValues[0]] &= $aParams_Split[$j] & '<br>' & @CRLF
				EndIf

				;Заменяем в описании параметра |True/1/False/0 на 1 и 0
				; Normalize markup in the parameter description.
				$aValues[$aValues[0]] = StringRegExpReplace($aValues[$aValues[0]], '(?i)\|(?:(.+))', '\1')
				$aValues[$aValues[0]] = StringRegExpReplace($aValues[$aValues[0]], '(?m)^\h*\+', '<br>')

				;Обрамляем в описании параметра [Opt*] тегами <b></b>
				; Highlight [Opt], [In], and [Out] markers using <b> tags.
				$aValues[$aValues[0]] = StringRegExpReplace($aValues[$aValues[0]], '(?i)(\[(?:Opt|In|Out)[^\]]*\])', '<b>\1</b>')

				; Assume any functions mentioned in parameters has an underscore either in the beginning or somewhere in the middle. Also avoid variables.
				$asRegExp = StringRegExp($aValues[$aValues[0]], "(?<!\$)\b\w*_\w+\b", $STR_REGEXPARRAYGLOBALMATCH)

				If IsArray($asRegExp) Then
					For $k = 0 To UBound($asRegExp) - 1
;~ 						ConsoleWrite($asRegExp[$k] & @CRLF)
						If __LinkIsValid($asRegExp[$k]) Then $aValues[$aValues[0]] = StringRegExpReplace($aValues[$aValues[0]], "(?<!\$)\b(\Q" & $asRegExp[$k] & "\E)\b(?!\.htm|</a>)", '<a href="\1.htm">\1</a>')
					Next
				EndIf

			Next

			ReDim $aParams[$aParams[0] + 1]
			ReDim $aValues[$aParams[0] + 1]
			#EndRegion Получаем и форматируем параметры функций ; Get and format function parameters

			#Region Получаем и форматируем возвращаемое значение функций ; Get and format function return values
			$aReturn = StringRegExp($aHeaders[$i][$__g_Header_Return], '(?si).*?Success\h*[-:]\h*(.*?)\s+Failure\s*[-:]\s*(.*)', $STR_REGEXPARRAYGLOBALMATCH)

			If @error Then
				$sReturn_Success = $aHeaders[$i][$__g_Header_Return]
				$sReturn_Failure = "None."
			Else
				$sReturn_Success = StringReplace($aReturn[0], @CRLF, '<br>' & @CRLF)
				$sReturn_Failure = StringReplace($aReturn[1], @CRLF, '<br>' & @CRLF)
			EndIf

			;Заменяем в описании возвр. значения |True/1/False/0 на 1 и 0.
			; Normalize markup in the return value description.
			$sReturn_Success = StringRegExpReplace($sReturn_Success, '(?i)\|\-', '')
			$sReturn_Success = StringRegExpReplace($sReturn_Success, '(?i)\|(?:(.+?))', '\1')
			$sReturn_Success = StringRegExpReplace($sReturn_Success, '<br>\h*\+', '')
			$sReturn_Failure = StringRegExpReplace($sReturn_Failure, '(?i)\|\-', '')
			$sReturn_Failure = StringRegExpReplace($sReturn_Failure, '(?i)\|(?:(.+?))', '\1')
			$sReturn_Failure = StringRegExpReplace($sReturn_Failure, '<br>\h*\+', '')

			$sReturn_Success = StringRegExpReplace($sReturn_Success, "(?i)((?|\@Error|\@Extended|Return):)", "<b>\1</b>")
			$sReturn_Failure = StringRegExpReplace($sReturn_Failure, "(?i)((?|\@Error|\@Extended):)", "<b>\1</b>")

			If StringRegExp($sReturn_Success, "^.*?[\r\n]{2}") Then
				$sReturn_Success &= '<br>&nbsp;'
			EndIf
			#EndRegion Получаем и форматируем возвращаемое значение функций ; Get and format function return values

			#Region  Формируем Html'ку функций' ; Build the function HTML page
			$sHtml_Content = _AU3Lib_GetHtmlHeaderStr($sDocsFunc_FormattedTmp_Title, '../css') & _
					'<body>' & @CRLF & '  <h1 class="small">' & $sDocsFuncHeader_FormattedTmp_Title & '</h1>' & @CRLF & _ ;~ modified
					'  <hr style="height:0px">' & @CRLF & _ ;~ modified/inserted
					'  <h1>' & $sFuncName & '</h1>' & @CRLF & _
					'  <p class="funcdesc">' & $aHeaders[$i][$__g_Header_Description] & '<br /></p>' & @CRLF & _
					'  <p class="codeheader">' & @CRLF & _
					$aHeaders[$i][$__g_Header_Syntax] & '<br>' & @CRLF & _
					'  </p>' & @CRLF & @CRLF & _
					'  <h2>Parameters</h2>' & @CRLF

			If $aParams[0] > 0 Then
				$sHtml_Content &= '  <table>' & @CRLF ;~ modified

				$sParam_TD_Width = ' style="width:15%"' ;~ modified
				$sValue_TD_Width = ' style="width:85%"' ;~ modified

				For $j = 1 To $aParams[0]
					If $j > 1 Then
						$sParam_TD_Width = ''
						$sValue_TD_Width = ''
					EndIf

					$sHtml_Content &= _
							'  <tr>' & @CRLF & _
							'    <td' & $sParam_TD_Width & '>' & $aParams[$j] & '</td>' & @CRLF & _
							'    <td' & $sValue_TD_Width & '>' & $aValues[$j] & '</td>' & @CRLF & _
							'  </tr>' & @CRLF
				Next

				$sHtml_Content &= '  </table>' & @CRLF & @CRLF
			Else
				$sHtml_Content &= 'None.<br>' & @CRLF
			EndIf

			$sHtml_Content &= '  <h2>Return Value</h2>' & @CRLF ;~ modified

			$sHtml_Content &= _
					'  <table class="noborder">' & @CRLF & _ ;~ modified
					'    <tr>' & @CRLF & _
					'      <td style="width:10%" class="valign-top">Success:</td>' & @CRLF & _
					'      <td style="width:90%">' & $sReturn_Success & '</td>' & @CRLF & _
					'    </tr>' & @CRLF & _
					'    <tr>' & @CRLF & _
					'      <td class="valign-top">Failure:</td>' & @CRLF & _
					'      <td>' & $sReturn_Failure & '</td>' & @CRLF & _
					'    </tr>' & @CRLF & _
					'  </table>' & @CRLF

			If $aHeaders[$i][$__g_Header_Remarks] <> '' Then ;~ modified
				$sRemarks = StringReplace($aHeaders[$i][$__g_Header_Remarks], @CRLF, '<br>' & @CRLF)
				$sRemarks = StringRegExpReplace($sRemarks, '(?m)^\h*\+', '&nbsp;') ;~ modified

				$sHtml_Content &= _
						@CRLF & '  <h2>Remarks</h2>' & @CRLF & _ ;~ modified
						$sRemarks & '<br>' & @CRLF & @CRLF ;~ modified
			EndIf ;~ modified

			If $aHeaders[$i][$__g_Header_Related] <> '' Then
				$aSplit = StringRegExp($aHeaders[$i][$__g_Header_Related], '(\w+)', $STR_REGEXPARRAYGLOBALMATCH)

				$aHeaders[$i][$__g_Header_Related] = ''

				For $j = 0 To UBound($aSplit) - 1
;~ 					$aHeaders[$i][$__g_Header_Related] &= ($aHeaders[$i][$__g_Header_Related] ? ', ' : '') & '<a href="' & $aSplit[$j] & '.htm">' & $aSplit[$j] & '</a>'
					If __LinkIsValid($aSplit[$j]) Then ; If function will be in the chm, add a link.
						$aHeaders[$i][$__g_Header_Related] &= ($aHeaders[$i][$__g_Header_Related] ? ', ' : '') & '<a href="' & $aSplit[$j] & '.htm">' & $aSplit[$j] & '</a>'

					Else ; Otherwise just add it to the list.
						$aHeaders[$i][$__g_Header_Related] &= ($aHeaders[$i][$__g_Header_Related] ? ', ' : '') & $aSplit[$j]

					EndIf
				Next

				$sRelated = StringReplace($aHeaders[$i][$__g_Header_Related], @CRLF, '<br>' & @CRLF)
				$sRelated = StringRegExpReplace($sRelated, '(?i)\|(?:(.+))', '\1')

				$sHtml_Content &= _
						'  <h2>Related</h2>' & @CRLF & _
						$sRelated & '<br>' & @CRLF
			EndIf

			If ($aHeaders[$i][$__g_Header_Example] <> "No") Then
				If (($aHeaders[$i][$__g_Header_Example] = "Yes") And FileExists($sLibExamplesDir & "\" & $sFile & ".au3")) Then ; Example named the same as the function.
					$sTempExamplePath = $sLibExamplesDir & "\"
					$sTempExampleName = $sFile

					ReDim $asExamples[1]
					$asExamples[0] = $sTempExamplePath & $sTempExampleName & ".au3"

					$hSearch = FileFindFirstFile($sTempExamplePath & $sTempExampleName & "*" & ".au3")

					If ($hSearch <> -1) Then  ; Find all additional examples
						Do
							$sFileName = FileFindNextFile($hSearch)
							If @error Then ExitLoop
							If StringRegExp($sFileName, "(?i)\Q" & $sTempExampleName & "\E\s*\[\d+\]\s*\.au3") Then ; Make sure this is an additional example, not just a similarly named one. Allow for spaces before/after the brackets ([]).
								ReDim $asExamples[UBound($asExamples) + 1]
								$asExamples[UBound($asExamples) - 1] = $sTempExamplePath & $sFileName
							EndIf
						Until @error
						FileClose($hSearch)
					EndIf

;~ 				ElseIf FileExists($aHeaders[$i][$__g_Header_Example]) Then ; Custom input path, or relative to @WorkingDir (Matches original functionality of MrCreatoR).
				ElseIf FileExists($aHeaders[$i][$__g_Header_Example]) Or FileExists($sLibraryDir & "\" & $aHeaders[$i][$__g_Header_Example]) Then ; Custom input path, or relative to Library directory (Matches original functionality of MrCreatoR).
;~ 					If FileExists(@WorkingDir & "\" & $aHeaders[$i][$__g_Header_Example]) Then; Example is in the working dir (The library source folder), or relative to it.
					If FileExists($sLibraryDir & "\" & $aHeaders[$i][$__g_Header_Example]) Then ; Example is in the working dir (The library source folder), or relative to it.
;~ 						$sTempExamplePath = @WorkingDir
						$sTempExamplePath = $sLibraryDir
						$sTempExampleName = $aHeaders[$i][$__g_Header_Example]

					Else ; This is a path to an example.
						$sTempExamplePath = StringMid($aHeaders[$i][$__g_Header_Example], 1, StringInStr($aHeaders[$i][$__g_Header_Example], "\", 0, -1))
						$sTempExampleName = StringMid($aHeaders[$i][$__g_Header_Example], StringInStr($aHeaders[$i][$__g_Header_Example], "\", 0, -1) + 1)
					EndIf

					ReDim $asExamples[1]
					$asExamples[0] = $sTempExamplePath & $sTempExampleName

				ElseIf FileExists($sLibExamplesDir & "\" & $aHeaders[$i][$__g_Header_Example]) Then ; Alternatively named example in the Example Dir.
					$sTempExamplePath = $sLibExamplesDir & "\"
					$sTempExampleName = $aHeaders[$i][$__g_Header_Example]

					ReDim $asExamples[1]
					$asExamples[0] = $sTempExamplePath & $sTempExampleName

				ElseIf FileExists($sLibExamplesDir & "\" & $aHeaders[$i][$__g_Header_Example] & ".au3") Then ; Alternatively named example in the Example Dir.
					$sTempExamplePath = $sLibExamplesDir & "\"
					$sTempExampleName = $aHeaders[$i][$__g_Header_Example]

					ReDim $asExamples[1]
					$asExamples[0] = $sTempExamplePath & $sTempExampleName & ".au3"
				EndIf
			EndIf

			If (UBound($asExamples) > 0) Then
				$sHtml_Content &= @CRLF & '  <h2>Example</h2>' & @CRLF & _ ;~ modified
						'  <script type="text/javascript">' & @CRLF & _
						'  if ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4)) // IE (4+) only' & @CRLF & _
						'  function copyToClipboard(section) {' & @CRLF & _
						'  if (window.clipboardData && clipboardData.setData) {' & @CRLF & _
						'  clipboardData.setData("text", section + "\r\n");' & @CRLF & _
						'  alert("Copied to clipboard");' & @CRLF & _
						'  }' & @CRLF & _
						'  }' & @CRLF & _
						'  </script>' & @CRLF

				For $iExample = 0 To UBound($asExamples) - 1
					If $iHighlightExampleSyntax Then

						$hTempFile = FileOpen($sTempFile, BitOR($FO_OVERWRITE, $FO_UTF8))
						FileWrite($hTempFile, FileRead($asExamples[$iExample]))
						FileClose($hTempFile)
						_SendSciTE_Command($hGUI, $hSciTE_hwnd, 'open:' & StringReplace($sTempFile, '\', '\\') & '') ; Open the temporary Au3 file.
						_SendSciTE_Command($hGUI, $hSciTE_hwnd, 'exportashtml:' & StringReplace($sTempOutputFile, '\', '\\')) ; Export it to html which will allow it to be colored.
						_SendSciTE_Command($hGUI, $hSciTE_hwnd, 'close:') ; Close the temporary Au3 file so I can write the new data to it.
						$sExample_Content = FileRead($sTempOutputFile)
						_SciTE_ParseHTML($sExample_Content) ; Content is ByRef modified.

						If $iHiglExmplSyntx_AddURLs Then $sExample_Content = __AddHelpLinks($sExample_Content)
					Else
						$sExample_Content = FileRead($asExamples[$iExample])
					EndIf

					$sExample_Content = StringStripWS($sExample_Content, $STR_STRIPLEADING + $STR_STRIPTRAILING)
					$sExample_Content = StringReplace($sExample_Content, '<>', '&lt;&gt;')

					If Not $iHighlightExampleSyntax Then
						$sExample_Content = StringReplace($sExample_Content, '<', '&lt;')
						$sExample_Content = StringReplace($sExample_Content, '>', '&gt;')
					EndIf

					If $sExample_Content <> "" Then
						If (UBound($asExamples) > 1) Then $sHtml_Content &= '  <h3>Example ' & ($iExample + 1) & '</h3>' & @CRLF
						$sHtml_Content &= _
								'  <div class="codeSnippetContainer">' & @CRLF & _
								'    <div class="codeSnippetContainerTabs">' & @CRLF & _
								'  <script type="text/javascript">' & @CRLF & _
								'  if (document.URL.match(/^mk:@MSITStore:/i)) {' & @CRLF & _
								'  document.write(''<div class="codeSnippetContainerTab codeSnippetContainerTabSingle" dir="ltr">'');' & @CRLF & _
								'  document.write(''<object id=hhctrl type="application/x-oleobject" classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11"><param name="Command" value="ShortCut"><param name="Font" value="Verdana,10pt"><param name="Text" value="Text:Open this Script"><param name="Item1" value=",Examples\\' & StringMid($asExamples[$iExample], StringInStr($asExamples[$iExample], "\", 0, -1) + 1) & ',"></object>'');' & @CRLF & _ ; $sFuncName & '.au3,"></object>'');' & @CRLF & _
								'  document.write(''<\/div>'');' & @CRLF & _
								'  }' & @CRLF & _
								'  </script>' & @CRLF & _
								'  </div>' & @CRLF & _
								' ' & @CRLF & _
								'  <div class="codeSnippetContainerCodeContainer">' & @CRLF & _
								'  <div class="codeSnippetToolBar">' & @CRLF & _
								'  <div class="codeSnippetToolBarText">' & @CRLF & _
								'  <script type="text/javascript">' & @CRLF & _
								'  if ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4)) // IE (4+) only' & @CRLF & _
								'  document.write(''<a href="#" id="copy" onclick="copyToClipboard(document.getElementById(\''copytext' & $iExample + 1 & '\'').innerText)">Copy to clipboard<\/a>'');' & @CRLF & _
								'  </script>' & @CRLF & _
								'  </div>' & @CRLF & _
								'  </div>' & @CRLF & _
								'  <div class="codeSnippetContainerCode" dir="ltr" id="copytext' & $iExample + 1 & '">' & @CRLF & _
								'  <pre>' & @CRLF & _
								$sExample_Content & @CRLF & _
								'  </pre>' & @CRLF & _
								'  </div>' & @CRLF & _
								'  </div>' & @CRLF & _
								'  </div>' & @CRLF
					EndIf
				Next
			EndIf

			;Закрываем теги Html'ки функций: Body, Html
			; Close the function HTML tags (Body, Html).
			$sHtml_Content &= '</body>' & @CRLF & '</html>'
			#EndRegion  Формируем Html'ку функций' ; Build the function HTML page

			;Пишем файл функций
			; Write the function HTML file.
			$hFunc_File = FileOpen($sDocsDir & "\funcs\" & $sFile & ".htm", $FO_OVERWRITE + $FO_UTF8)
			FileWrite($hFunc_File, $sHtml_Content)
			FileClose($hFunc_File)
		Next

		; Close the Sub-Component Index table, body, and html tags.
		FileWriteLine($hSubCompIndex, "  </table>" & @CRLF & "</body>" & @CRLF & "</html>")
		FileClose($hSubCompIndex)
	Next

	;Закрываем теги: table, Body, Html
	; Close the Index body, and html tags.
	FileWriteLine($hIndex_File, "</body>" & @CRLF & "</html>")

	;Закрываем файл индекса
	; Close the index file.
	FileClose($hIndex_File)

	If $iHighlightExampleSyntax Then
		FileClose($hTempFile)
		FileDelete($sTempFile)
		FileDelete($sTempOutputFile)
		If StringInStr($sCurrentFile, "\") Then _SendSciTE_Command($hGUI, $hSciTE_hwnd, 'open:' & StringReplace($sCurrentFile, '\', '\\') & '') ; Restore Scite to the current tab.
	EndIf

	For $i = 0 To UBound($avIndexFiles) - 1
		; Close the Component Index files table, body, and html tags.
		FileWriteLine($avIndexFiles[$i][$__eIndexHandle], "  </table>" & @CRLF & "</body>" & @CRLF & "</html>")

		; Close the Component index file.
		FileClose($avIndexFiles[$i][$__eIndexHandle])

		$avIndexFiles[$i][$__eIndexHandle] = Null
	Next

	; Replace the generated Index.htm with a user defined Index.htm.
	; If <LibraryName>_Index.htm (library specific Index) exists in the AU3 Library File Directory then use this file as Index.htm
	; If <LibraryName>_Index.htm does not exist but a general Index.htm exists in the AU3 Library File Directory then use this file as Index.htm
	; If none of both files exists then the generated Index.htm is used.
	Local $sLibraryPath = StringLeft($sLibraryFile, StringInStr($sLibraryFile, "\", 0, -1))
	If FileExists($sLibraryPath & $sLibraryFile_Name & "_Index.htm") Then
		FileCopy($sLibraryPath & $sLibraryFile_Name & "_Index.htm", $sIndex_File, 1) ;~ modified MODIFY!!
	ElseIf FileExists($sLibraryPath & "Index.htm") Then
		FileCopy($sLibraryPath & "Index.htm", $sIndex_File, 1) ;~ modified MODIFY!!
	EndIf

	Return $sIndex_File
EndFunc   ;==>_AU3Lib_GenerateDocs_Proc

;$aHeaders[0][1-10] = Header key name
;$aHeaders[N][0] = Name
;$aHeaders[N][1] = Description
;$aHeaders[N][2] = Syntax
;$aHeaders[N][3] = Parameters
;$aHeaders[N][4] = Return values
;$aHeaders[N][5] = Author
;$aHeaders[N][6] = Modified
;$aHeaders[N][7] = Remarks
;$aHeaders[N][8] = Related
;$aHeaders[N][9] = Link
;$aHeaders[N][10] = Example
Func _AU3Lib_GetHeaders($sLibraryFile)
	If Not FileExists($sLibraryFile) Then Return SetError(1, 0, -1)
	Local $aLibHeaders = StringRegExp(FileRead($sLibraryFile), '(?s).*?(; #FUNCTION# =+.*?; =+).*?', $STR_REGEXPARRAYGLOBALMATCH)
	Local $iUbound = UBound($aLibHeaders) - 1
	Local $aHeaders[500][11], $aHeader_Params, $iCount, $sParam

	$aHeader_Params = _
			StringSplit( _
			'Name...........|Description....|Syntax.........|Parameters.....|Return values..|' & _
			'Author.........|Modified.......|Remarks........|Related........|Link...........|Example........', _
			'|')

	For $i = 0 To $iUbound
		If StringIsSpace($aLibHeaders[$i]) Then ContinueLoop

		$iCount += 1

		For $j = 1 To $aHeader_Params[0]
			$sParam = StringRegExpReplace($aLibHeaders[$i], '(?s).*; ' & $aHeader_Params[$j] & ': ?(.*?)[\r\n]; ([\w\s_]+\.+:|=+).*', '\1')
			If @extended = 0 Or $sParam = $aLibHeaders[$i] Then ContinueLoop

			$sParam = StringRegExpReplace($sParam, '(?m)^;', '')
			If StringIsSpace($sParam) Then $sParam = ''

			$aHeaders[0][$j - 1] = $aHeader_Params[$j]
			$aHeaders[$iCount][$j - 1] = StringStripWS($sParam, 3)
		Next
	Next

	ReDim $aHeaders[$iCount + 1][11]
	Return $aHeaders
EndFunc   ;==>_AU3Lib_GetHeaders

Func _AU3Lib_GetHtmlHeaderStr($sTitle, $sCss_Path = 'css')
	Return _
			'<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN">' & @CRLF & _
			'<html>' & @CRLF & _
			'<head>' & @CRLF & _
			'  <title>' & $sTitle & '</title>' & @CRLF & _
			'  <meta charset="UTF-8">' & @CRLF & _ ;~ modified
			'  <link href="' & $sCss_Path & '/default.css" rel="stylesheet">' & @CRLF & _ ;~ modified
			'</head>' & @CRLF
EndFunc   ;==>_AU3Lib_GetHtmlHeaderStr

Func _AU3Lib_GetValidFileName($sString, $sPatern = '[*?\\/|:<>"]', $sReplace = '_')
	If StringIsSpace($sString) Then Return $sString

	$sString = StringRegExpReplace($sString, $sPatern, $sReplace)
	Return SetExtended(@extended, StringRegExpReplace($sString, '(' & $sReplace & '+)', $sReplace))
EndFunc   ;==>_AU3Lib_GetValidFileName

Func _AU3Lib_CompileHtmlToChm($sChm_File_Path, $sChm_File_Name)
	GUICtrlSetData($idStatus_Label, "Please wait..." & @CRLF & "[Compiling to Chm]")

	Local $sPathHTMHelpCompiler = @ScriptDir & "\Resources"

	EnvSet("PATH", $sPathHTMHelpCompiler & "\hhc.exe;" & $sPathHTMHelpCompiler & "\hha.dll")
	RunWait($sPathHTMHelpCompiler & '\regsvr32.exe /s "' & $sPathHTMHelpCompiler & '\itcc.dll"')

	Local $hFile, $sIndex_hhk_File_Content, $sTOC_hhc_File_Content, $sChmProject_hhp_File_Content ; $aFileList,

	Local $sIndex_hhk_File = $sChm_File_Path & "\Index.hhk"
	Local $sTOC_hhc_File = $sChm_File_Path & "\TOC.hhc"
	Local $sChmProject_hhp_File = $sChm_File_Path & "\" & $sChm_File_Name & ".hhp"
	Local $sChm_File = $sChm_File_Path & "\" & $sChm_File_Name & ".chm"

	$sIndex_hhk_File_Content = _
			'<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">' & @CRLF & _
			'<HTML>' & @CRLF & _
			'<HEAD>' & @CRLF & _
			'<meta name="GENERATOR" content="Microsoft&reg; HTML Help Workshop 4.1">' & @CRLF & _
			'<!-- Sitemap 1.0 -->' & @CRLF & _
			'</HEAD><BODY>' & @CRLF & _
			'<UL>' & @CRLF

	$sTOC_hhc_File_Content = _
			'<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">' & @CRLF & _
			'<HTML>' & @CRLF & _
			'<HEAD>' & @CRLF & _
			'<meta name="GENERATOR" content="Microsoft&reg; HTML Help Workshop 4.1">' & @CRLF & _
			'<!-- Sitemap 1.0 -->' & @CRLF & _
			'</HEAD><BODY>' & @CRLF & _
			'<OBJECT type="text/site properties">' & @CRLF & _
			'  <param name="ImageType" value="Folder">' & @CRLF & _
			'</OBJECT>' & @CRLF & _
			'<UL>' & @CRLF & _
			'  <LI> <OBJECT type="text/sitemap">' & @CRLF & _
			'    <param name="Name" value="' & $sDocsIndexHeader_Title & '">' & @CRLF & _
			'    <param name="Local" value="index.htm">' & @CRLF & _
			'    </OBJECT>' & @CRLF & _
			'  <UL>' & @CRLF

	$sChmProject_hhp_File_Content = _
			'[OPTIONS]' & @CRLF & _
			'Compatibility=1.1 or later' & @CRLF & _
			'Compiled file=' & $sChm_File_Name & '.chm' & @CRLF & _
			'Contents file=TOC.hhc' & @CRLF & _
			'Default Window=Main' & @CRLF & _
			'Default topic=index.htm' & @CRLF & _
			'Display compile progress=Yes' & @CRLF & _
			'Full-text search=Yes' & @CRLF & _
			'Index file=Index.hhk' & @CRLF & _
			'Language=0x409 English (United States)' & @CRLF & _
			'Title=' & $sDocsIndex_Title & @CRLF & _
			'' & @CRLF & _
			'[WINDOWS]' & @CRLF & _
			'Main=,"TOC.hhc","Index.hhk","index.htm",,,,,,0x63520,,0x280e,[100,100,1020,800],,,,,,,0' & @CRLF & _
			'' & @CRLF & _
			'[FILES]' & @CRLF

	Local $asComponentIndex[0], $asSubComponentIndex[0], $asFuncs[0]

	; Get all components listed in the index.
	$asComponentIndex = StringRegExp(FileRead($sDocsDir & "\index.htm"), '<li><a href=.+?>(\w+)</a></li>', $STR_REGEXPARRAYGLOBALMATCH)
	; <li><a href="indices/Base_Index.htm">Base</a></li>
	If Not IsArray($asComponentIndex) Then
		ConsoleWrite("! No indexes found" & @CRLF)
		Return SetError(1, 0, 0)
	EndIf

	For $i = 0 To UBound($asComponentIndex) - 1
		If Not FileExists($sDocsDir & "\indices\" & $asComponentIndex[$i] & "_Index.htm") Then
			ConsoleWrite("! Component index file not found: " & $asComponentIndex[$i] & @CRLF)
			ContinueLoop
		EndIf

		; Add the Component to the TOC
		$sTOC_hhc_File_Content &= _
				'    <LI> <OBJECT type="text/sitemap">' & @CRLF & _
				'      <param name="Name" value="' & $asComponentIndex[$i] & '">' & @CRLF & _
				'      <param name="Local" value="indices\' & $asComponentIndex[$i] & '_Index.htm">' & @CRLF & _
				'    </OBJECT>' & @CRLF & _
				'    <UL>' & @CRLF

		; Add Component Index to the Index
		$sIndex_hhk_File_Content &= _
				'  <LI> <OBJECT type="text/sitemap">' & @CRLF & _
				'    <param name="Name" value="' & $asComponentIndex[$i] & '">' & @CRLF & _
				'    <param name="Local" value="indices\' & $asComponentIndex[$i] & '_Index.htm">' & @CRLF & _
				'  </OBJECT>' & @CRLF

		; Add Component Index to the CHM file
		$sChmProject_hhp_File_Content &= 'indices\' & $asComponentIndex[$i] & '_Index.htm' & @CRLF

		; Get a list of all Sub-Component indexes
		$asSubComponentIndex = StringRegExp(FileRead($sDocsDir & "\indices\" & $asComponentIndex[$i] & "_Index.htm"), '<td><a href=.+?>(\w+)</a></td>', $STR_REGEXPARRAYGLOBALMATCH)
		If Not IsArray($asSubComponentIndex) Then
			ConsoleWrite("! No sub-indexes found: " & $asComponentIndex[$i] & @CRLF)
			ContinueLoop
		EndIf

		For $j = 0 To UBound($asSubComponentIndex) - 1
			If Not FileExists($sDocsDir & "\indices\" & $asComponentIndex[$i] & "_" & $asSubComponentIndex[$j] & "_Index.htm") Then
				ConsoleWrite("! Sub-Component index file not found: " & $asSubComponentIndex[$j] & @CRLF)
				ContinueLoop
			EndIf

			; Add the Sub-Component to the TOC
			$sTOC_hhc_File_Content &= _
					'      <LI> <OBJECT type="text/sitemap">' & @CRLF & _
					'        <param name="Name" value="' & $asSubComponentIndex[$j] & '">' & @CRLF & _
					'        <param name="Local" value="indices\' & $asComponentIndex[$i] & "_" & $asSubComponentIndex[$j] & '_Index.htm">' & @CRLF & _
					'      </OBJECT>' & @CRLF & _
					'      <UL>' & @CRLF

			; Get a list of all Functions
			$asFuncs = StringRegExp(FileRead($sDocsDir & "\indices\" & $asComponentIndex[$i] & "_" & $asSubComponentIndex[$j] & "_Index.htm"), '<td><a href=.+?>(\w+)</a></td>', $STR_REGEXPARRAYGLOBALMATCH)
			If Not IsArray($asFuncs) Then
				ConsoleWrite("! No Functions found: " & $asSubComponentIndex[$j] & @CRLF)
				ContinueLoop
			EndIf

			For $k = 0 To UBound($asFuncs) - 1
				If Not FileExists($sDocsDir & "\funcs\" & $asFuncs[$k] & ".htm") Then
					ConsoleWrite("! Function file not found: " & $asFuncs[$k] & @CRLF)
					ContinueLoop
				EndIf

				; Add the Function to the TOC
				$sTOC_hhc_File_Content &= _
						'        <LI> <OBJECT type="text/sitemap">' & @CRLF & _
						'          <param name="Name" value="' & $asFuncs[$k] & '">' & @CRLF & _
						'          <param name="Local" value="funcs\' & $asFuncs[$k] & '.htm">' & @CRLF & _
						'        </OBJECT>' & @CRLF

				; Add function to the Index
				$sIndex_hhk_File_Content &= _
						'  <LI> <OBJECT type="text/sitemap">' & @CRLF & _
						'    <param name="Name" value="' & $asFuncs[$k] & '">' & @CRLF & _
						'    <param name="Local" value="funcs\' & $asFuncs[$k] & '.htm">' & @CRLF & _
						'  </OBJECT>' & @CRLF

				; Add function to the CHM file
				$sChmProject_hhp_File_Content &= 'funcs\' & $asFuncs[$k] & ".htm" & @CRLF
			Next

			; close off the Function UL.
			$sTOC_hhc_File_Content &= '      </UL>' & @CRLF

		Next
		; close off the sub-component UL etc.
		$sTOC_hhc_File_Content &= '    </UL>' & @CRLF

	Next

	$sIndex_hhk_File_Content &= _
			'</UL>' & @CRLF & _
			'</BODY></HTML>' & @CRLF

	$sTOC_hhc_File_Content &= _
			'  </UL>' & @CRLF & _
			'</UL>' & @CRLF & _
			'</BODY></HTML>' & @CRLF

	$sChmProject_hhp_File_Content &= _
			'css\h1_background.jpg' & @CRLF & _
			'' & @CRLF & _
			'[INFOTYPES]' & @CRLF & _
			'' & @CRLF

	$hFile = FileOpen($sIndex_hhk_File, $FO_OVERWRITE)
	FileWrite($hFile, $sIndex_hhk_File_Content)
	FileClose($hFile)

	; Replace the generated TOC.hhc with a user defined TOC.hhc.
	; If <LibraryName>_TOC.hhc (library specific TOC) exists in the AU3 Library File Directory then use this file as TOC.hhc
	; If <LibraryName>_TOC.hhc does not exist but a general TOC.hhc exists in the AU3 Library File Directory then use this file as TOC.hhc
	; If none of both files exists then the generated TOC.hhc is used.
	Local $sLibraryPath = StringLeft($sLibraryFile, StringInStr($sLibraryFile, "\", 0, -1))
	If FileExists($sLibraryPath & $sLibraryFile_Name & "_TOC.hhc") Then
		FileCopy($sLibraryPath & $sLibraryFile_Name & "_TOC.hhc", $sTOC_hhc_File) ;~ modified MODIFY!!
	ElseIf FileExists($sLibraryPath & "TOC.hhc") Then
		FileCopy($sLibraryPath & "TOC.hhc", $sTOC_hhc_File) ;~ modified MODIFY!!
	Else
		$hFile = FileOpen($sTOC_hhc_File, $FO_OVERWRITE)
		FileWrite($hFile, $sTOC_hhc_File_Content)
		FileClose($hFile)
	EndIf

	$hFile = FileOpen($sChmProject_hhp_File, $FO_OVERWRITE)
	FileWrite($hFile, $sChmProject_hhp_File_Content)
	FileClose($hFile)

	FileDelete($sChm_File)
	RunWait($sPathHTMHelpCompiler & '\hhc.exe "' & $sChm_File & '"', $sChm_File_Path, @SW_HIDE)
EndFunc   ;==>_AU3Lib_CompileHtmlToChm
#EndRegion Core Functions

#Region Internal Functions
Func _SaveSettings()
	; Save settings/data.
	IniWrite($sAppSettingFile, "Settings", "Library File", GUICtrlRead($idLibraryFile_Input))
	IniWrite($sAppSettingFile, "Settings", "Library Examples Dir", GUICtrlRead($idLibExamplesDir_Input))
	IniWrite($sAppSettingFile, "Settings", "Docs Dir", GUICtrlRead($idDocsDir_Input))
	IniWrite($sAppSettingFile, "Settings", "Docs Index Title", GUICtrlRead($idDocsIndexTitle_Input))
	IniWrite($sAppSettingFile, "Settings", "Docs Index Header Title", GUICtrlRead($idDocsIndexHeaderTitle_Input))
	IniWrite($sAppSettingFile, "Settings", "Docs Func Title", GUICtrlRead($idDocsFuncTitle_Input))
	IniWrite($sAppSettingFile, "Settings", "Docs Func Header Title", GUICtrlRead($idDocsFuncHeaderTitle_Input))
	IniWrite($sAppSettingFile, "Settings", "CHM File Name", GUICtrlRead($idCHMFileName_Input))
	IniWrite($sAppSettingFile, "Settings", "Show Generated Docs", GUICtrlRead($idShowGeneratedDocs_CB))
	IniWrite($sAppSettingFile, "Settings", "Highlight Example Syntax", GUICtrlRead($idHighlightExampleSyntax_CB))
	IniWrite($sAppSettingFile, "Settings", "Highlight Example Syntax Add URLs", GUICtrlRead($idHiglExmplSyntx_AddURLs_CB))
	IniWrite($sAppSettingFile, "Settings", "Compile To Chm", GUICtrlRead($idCompileToChm_CB))
	IniWrite($sAppSettingFile, "Settings", "Delete Source on Compile", GUICtrlRead($idCompile_DelOnDone_CB))
EndFunc   ;==>_SaveSettings

#EndRegion Internal Functions

; The below function was found and modified from AutoIt Help file maker, autoit-docs-v3.3.18.0-src, in file named: "SciteLib.au3". Credits to Jos, and AutoIt team.
Func __AddHelpLinks($sData)
	Local Const $sAutoIt_Link = "http://www.autoitscript.com/autoit3/docs"
	Local $asRegExp

	; Add Links to known keywords
	; LIBFUNC-structure of $tag ... in the examples
	$sData = StringRegExpReplace($sData, '<span class="S9">(\$tag\w+?)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/libfunctions/\1.htm"><span class="S9">\1</span></a>')
	; Functions
	$sData = StringRegExpReplace($sData, '<span class="S4">([\w]+?)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/functions/\1.htm"><span class="S4">\1</span></a>')
	$sData = StringReplace($sData, 'href="/functions/Opt.htm">', 'href="/functions/AutoItSetOption.htm">')
	; Exception for UDPStartup(), UDPShutdown()
	$sData = StringReplace($sData, '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/functions/UDPStartup.htm"><span class="S4">UDPStartup</span></a>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/functions/TCPStartup.htm"><span class="S4">UDPStartup</span></a>')
	$sData = StringReplace($sData, '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/functions/UDPShutdown.htm"><span class="S4">UDPShutdown</span></a>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/functions/TCPShutdown.htm"><span class="S4">UDPShutdown</span></a>')

	; Macros
	$sData = StringRegExpReplace($sData, '(?i)<span class="S6">(@[^<]+)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/macros.htm#\1"><span class="S6">\1</span></a>')
	; Operators
	$sData = StringRegExpReplace($sData, '(?i)<span class="S8">((?:[+^*/=&-]|&gt;|&lt;)+)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/intro/lang_operators.htm"><span class="S8">\1</span></a>')
	; Keywords
	$sData = StringRegExpReplace($sData, '(?i)<span class="S5">(Not|And|Or)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/intro/lang_operators.htm"><span class="S5">\1</span></a>')
	$sData = StringRegExpReplace($sData, '(?i)<span class="S5">(ContinueCase|ContinueLoop|Default|Dim|Do|Enum|Exit|ExitLoop|For|Func|If|ReDim|Select|Static|Switch|While|With)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/\1.htm"><span class="S5">\1</span></a>')
	$sData = StringRegExpReplace($sData, '(?i)<span class="S5">(Else|Then|ElseIf|EndIf)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/IfElseEndIf.htm"><span class="S5">\1</span></a>')
	$sData = StringRegExpReplace($sData, '(?i)<span class="S5">(Next|To|Step)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/For.htm"><span class="S5">\1</span></a>')
	$sData = StringRegExpReplace($sData, '(?i)<span class="S5">(Case|EndSwitch)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/Switch.htm"><span class="S5">\1</span></a>')
	$sData = StringRegExpReplace($sData, '(?i)<span class="S5">(Global|Local|Const)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/Dim.htm"><span class="S5">\1</span></a>')
	$sData = StringRegExpReplace($sData, '(?i)<span class="S5">(EndFunc|ByRef|Return)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/Func.htm"><span class="S5">\1</span></a>')
	$sData = StringRegExpReplace($sData, '(?i)<span class="S5">(True|False)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/Booleans.htm"><span class="S5">\1</span></a>')
	$sData = StringReplace($sData, '<span class="S5">Until</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/Do.htm"><span class="S5">Until</span></a>')
	$sData = StringReplace($sData, '<span class="S5">WEnd</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/While.htm"><span class="S5">WEnd</span></a>')
	$sData = StringReplace($sData, '<span class="S5">EndSelect</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/Select.htm"><span class="S5">EndSelect</span></a>')
	$sData = StringReplace($sData, '<span class="S5">In</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/ForInNext.htm"><span class="S5">In</span></a>')
	$sData = StringReplace($sData, '<span class="S5">EndWith</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/With.htm"><span class="S5">EndWith</span></a>')
	; Directives
	$sData = StringReplace($sData, '<span class="S11">#OnAutoItStartRegister</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/OnAutoItStartRegister.htm"><span class="S11">#OnAutoItStartRegister</span></a>')
	$sData = StringReplace($sData, '<span class="S11">#include</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/include.htm"><span class="S11">#include</span></a>')
	$sData = StringReplace($sData, '<span class="S11">#include-once</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/include-once.htm"><span class="S11">#include-once</span></a>')
	$sData = StringReplace($sData, '<span class="S11">#RequireAdmin</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/RequireAdmin.htm"><span class="S11">#RequireAdmin</span></a>')
	$sData = StringReplace($sData, '<span class="S11">#NoTrayIcon</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/keywords/NoTrayIcon.htm"><span class="S11">#NoTrayIcon</span></a>')

	; Add links to known UDFs. Scite only exports built-in UDFs as S15.
	$sData = StringRegExpReplace($sData, '<span class="S15">([\w]+?)</span>', '<a class="codeSnippetLink" href="' & $sAutoIt_Link & '/libfunctions/\1.htm"><span class="S15">\1</span></a>')

	; Obtain a list of in-file UDFs
	$asRegExp = StringRegExp($sData, '\Q<a class="codeSnippetLink" href="http://www.autoitscript.com/autoit3/docs/keywords/Func.htm"><span class="S5">Func</span></a>\E<span class="S0">[ &;\w]*?([\w]+?)</span>', $STR_REGEXPARRAYGLOBALMATCH)

	If IsArray($asRegExp) Then
		For $i = 0 To UBound($asRegExp) - 1
			; Set in-file UDFs as style 17, like Scite does.
			$sData = StringRegExpReplace($sData, '(<span class="S0">[ &;\w]*?)(' & $asRegExp[$i] & ')</span>', '\1</span><span class="S17">\2</span>')
		Next
	EndIf

	; Obtain a list of remaining un-linked, uncolored potential UDFs colored as S0 (Whitespace), Scite exports it as that for UDFs in complex includes (More than 1 layer deep).
	$asRegExp = StringRegExp($sData, '<span class="S0">[\s&;\w]*?([\w]+?)</span>', $STR_REGEXPARRAYGLOBALMATCH)

	If IsArray($asRegExp) Then
		For $i = 0 To UBound($asRegExp) - 1
			If FileExists($sLibExamplesDir & "\" & $asRegExp[$i] & ".au3") Then

				; Only set to Style 17, because Scite might have styled some already.
				$sData = StringRegExpReplace($sData, '(<span class="S0">[ &;\w]*?\s)(' & $asRegExp[$i] & ')</span>', '\1</span><span class="S17">\2</span>')
			EndIf
		Next
	EndIf

	; Obtain a list of remaining un-linked, UDFs colored as S17.
	$asRegExp = StringRegExp($sData, '<span class="S17">([\w]+?)</span>', $STR_REGEXPARRAYGLOBALMATCH)

	If IsArray($asRegExp) Then
		For $i = 0 To UBound($asRegExp) - 1
			; Only add links to valid functions included in this help file.
;~ 			For $j = 0 To UBound($asFuncNames) - 1
;~ 				If StringRegExp($asFuncNames[$j], "\b" & $asRegExp[$i] & "\b") Then
			If __LinkIsValid($asRegExp[$i]) Then
				$sData = StringRegExpReplace($sData, '<span class="S17">(' & $asRegExp[$i] & ')</span>', '<a class="codeSnippetLink" href="\1.htm"><span class="S17">\1</span></a>')
;~ 					ExitLoop
			EndIf
;~ 			Next
		Next
	EndIf

	Return $sData
EndFunc   ;==>__AddHelpLinks

; The below function was found and modified from AutoIt Help file maker, autoit-docs-v3.3.18.0-src, in file named: "SciteLib.au3". Credits to Jos, and AutoIt team.
Func _SciTE_ParseHTML(ByRef $sData)
	Local $bReturn = False
	; Parse the HTML output.
	$sData = StringRegExp($sData & @CRLF & _
			'<body bgcolor="#000000">' & @CRLF & _
			'' & @CRLF & _
			'</body>' & @CRLF, '(?is:<body bgcolor="#[[:xdigit:]]{6}">\R(.*?)\R</body>\R)', $STR_REGEXPARRAYGLOBALMATCH)[0] ; Access the first element, as this will always return something due to having a fallback.
	$sData = StringReplace($sData, '<br />', '')
	; If not blank, then return True.
	$bReturn = Not (StringStripWS($sData, $STR_STRIPALL) == '')
	Return $bReturn
EndFunc   ;==>_SciTE_ParseHTML

Func __LinkIsValid($sFuncName, $asFiles = Null) ; Parses and stores a list of all functions that will be processed, so I can know which functions will be valid links.
	Local $asRegExp
	Local Static $asFuncNames[0]
	Local $iCount = 0

	If ($sFuncName = Default) And IsArray($asFiles) Then ; Parse and store all function names, overwriting the last ones stored, if any.
		ReDim $asFuncNames[UBound($asFiles)]

		For $iSrcFile = 0 To UBound($asFiles) - 1
			If ($asFiles[$iSrcFile] = "") Then ContinueLoop

			$asRegExp = StringRegExp(FileRead($asFiles[$iSrcFile]), "(?i);\s*Name\s*\.+:\s*(?!__)(\w+)", $STR_REGEXPARRAYGLOBALMATCH)

			If IsArray($asRegExp) Then
				$asFuncNames[$iCount] = " "
				For $i = 0 To UBound($asRegExp) - 1
					$asFuncNames[$iCount] &= $asRegExp[$i] & " "
				Next
				$iCount += 1
			EndIf
		Next

		ReDim $asFuncNames[$iCount]

		Return True

	ElseIf IsString($sFuncName) Then
		For $i = 0 To UBound($asFuncNames) - 1
			If StringRegExp($asFuncNames[$i], "\b" & $sFuncName & "\b") Then Return True
		Next
	EndIf

	Return False
EndFunc   ;==>__LinkIsValid

#Region ; The following Scite functions are copied and modified from AutoItWrapper. Credits to Jos.

; Receive Data from SciTE
Func _MY_WM_COPYDATA($HWindow, $msg, $wParam, $lParam)
	#forceref $HWindow, $msg, $wParam
	Local $COPYDATA = DllStructCreate('Ptr;DWord;Ptr', $lParam)
	Local $SciTECmdLen = DllStructGetData($COPYDATA, 2)
	Local $CmdStruct = DllStructCreate('Char[255]', DllStructGetData($COPYDATA, 3))
	$sSciTECmd = StringLeft(DllStructGetData($CmdStruct, 1), $SciTECmdLen)
;~ 	ConsoleWrite('<--' & $sSciTECmd & @CRLF)
EndFunc   ;==>_MY_WM_COPYDATA

Func _SendSciTE_Command($My_Hwnd, $hSciTE_hwnd, $sCmd)
	Local $CmdStruct = DllStructCreate('Char[' & StringLen($sCmd) + 1 & ']')
	DllStructSetData($CmdStruct, 1, $sCmd)
	Local $COPYDATA = DllStructCreate('Ptr;DWord;Ptr')
	DllStructSetData($COPYDATA, 1, 1)
	DllStructSetData($COPYDATA, 2, StringLen($sCmd) + 1)
	DllStructSetData($COPYDATA, 3, DllStructGetPtr($CmdStruct))
	DllCall('User32.dll', 'None', 'SendMessageA', 'HWnd', $hSciTE_hwnd, _
			'Int', $WM_COPYDATA, 'HWnd', $My_Hwnd, _
			'Ptr', DllStructGetPtr($COPYDATA))
;~ __ConsoleWrite('-->' & $sCmd & @CRLF)
EndFunc   ;==>_SendSciTE_Command
;
Func _SendSciTE_GetInfo($My_Hwnd, $hSciTE_hwnd, $sCmd)
	Local $My_Dec_Hwnd = Dec(StringTrimLeft($My_Hwnd, 2))
	$sCmd = ":" & $My_Dec_Hwnd & ":" & $sCmd
	$sSciTECmd = ""
	_SendSciTE_Command($My_Hwnd, $hSciTE_hwnd, $sCmd)
	For $x = 1 To 10
		If $sSciTECmd <> "" Then ExitLoop
		Sleep(20)
	Next
	$sSciTECmd = StringTrimLeft($sSciTECmd, StringLen(":" & $My_Dec_Hwnd & ":"))
	$sSciTECmd = StringReplace($sSciTECmd, "macro:stringinfo:", "")
	Return $sSciTECmd
EndFunc   ;==>_SendSciTE_GetInfo
#EndRegion ; The following Scite functions are copied and modified from AutoItWrapper. Credits to Jos.
