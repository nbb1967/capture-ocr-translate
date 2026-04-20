#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ico\cot-app.ico
#AutoIt3Wrapper_Outfile=COT.exe
#AutoIt3Wrapper_Res_Comment=Utility for recognizing and translating text captured from the screen...
#AutoIt3Wrapper_Res_Description=Capture-OCR-Translate
#AutoIt3Wrapper_Res_Fileversion=0.9.6.113
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=Capture-OCR-Translate
#AutoIt3Wrapper_Res_ProductVersion=0.9.6
#AutoIt3Wrapper_Res_CompanyName=NyBumBum
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © NyBumBum 2025–2026
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Icon_Add=ico\white.ico
#AutoIt3Wrapper_Res_Icon_Add=ico\black.ico
#AutoIt3Wrapper_Res_Icon_Add=ico\white-active.ico
#AutoIt3Wrapper_Res_Icon_Add=ico\black-active.ico
#AutoIt3Wrapper_Res_File_Add=loc\en.ini, 6
#AutoIt3Wrapper_Res_File_Add=loc\ru.ini, 6
#AutoIt3Wrapper_Res_File_Add=cur\cross.cur, 1, 101
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


;Hey! Don't forget to replace the cursor in Resource Hacker....
;Don't forget to change the YEAR in the copyright ABOVE and in the About tab...


#include <Array.au3>
#include <ButtonConstants.au3>
#include <Clipboard.au3>
#include <ComboConstants.au3>
#include <Constants.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <GDIPlus.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiEdit.au3>
#include <GuiMenu.au3>
#include <Misc.au3>
#include <ScreenCapture.au3>
#include <StaticConstants.au3>
#include <TabConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#include <WinAPILocale.au3>
#include <WinAPIRes.au3>
#include <WinAPISys.au3>
#include <WinAPISysWin.au3>
#include "include\GuiHotkey.au3" ; UDF by Mat
#include "include\JSON.au3"      ; UDF by AspirinJunkie    https://github.com/Sylvan86/autoit-json-udf

;-----------------------------------------------------------------CHECKING IF ALREADY RUNNING
if _Singleton("Capture_OCR_Translate", 1) = 0 Then
    WinActivate("[CLASS:AutoIt v3 GUI;TITLE:Capture-OCR-Translate]")
	Exit
EndIf
;----------------------
Opt("GUICloseOnESC", 0)
Opt("TrayMenuMode", 1+2+4)

Global $bIsProcessing = False ; flag to prevent the creation of a task queue

Global Const $sRegexCurlError = '(curl:\s\([0-9]+\).*)'
Global $bNeedAPIKey = False
Global $idSelectAll	; for castom menu
Global $hEdit, $hInput
Global $hCursorArrow, $hMyCursor
Global $sSystemProxy = _GetSystemProxyForCurl()
Global $sGoogleError, $sGoogleErrorCode

Global $sPng_Path = @TempDir & "\tmp.png"
Global $sJpg_Path = @TempDir & "\tmp.jpg"

Global $nAppsUseLightTheme, $nSystemUsesLightTheme
_Detect_Dark_Light_Theme()

Global $hInstance = _WinAPI_GetModuleHandle(0)
If $hInstance = 0 Then
	MsgBox($MB_ICONERROR, "Error", "Failed to get program handle.")
	Exit
EndIf

Local $hDLL = DllOpen("user32.dll"); for _IsPressed

#Region;===================== LOCALIZATION (BASED ON FAQ BY YASHIED) ========================
;----------------------------------------------EXTRACT STRINGS FROM A STRING TABLE
Func _GetStrRes($iStringID)
	Local $iString = _WinAPI_LoadString($hInstance, $iStringID)
	If @error Then
		MsgBox($MB_ICONERROR, "Error", "Failed to get string from program resources.")
		Exit
	EndIf
	Return $iString
EndFunc   ;==>_GetStrRes
;------------------------------------
Global Enum $eAppName, $eTabResult, $eTabGeneral, $eTabAPIKey, $eTabHotkey, $eTabLanguage, $eTabAbout
Global $asArrayGUI_App[7] = [6000, 6001, 6002, 6003, 6004, 6005, 6006]
Global Enum $eClipboard, $eAlwaysOnTop, $eTrayIconAction
Global $asArrayTabGeneral[3] = [6016, 6017, 6018]
;----------
Global $asArrayTrayIconAction[3] = [6032, 6033, 6034]
Global Enum $eDescription, $eGetOCRAPIKey, $eSave
Global $asArrayTabAPIKey[3] = [6048, 6049, 6050]
Global Enum $eOCRHotkey, $eTransHotkey, $eTipHotkey
Global $asArrayTabHotkey[3] = [6064, 6065, 6066]
Global Enum $eTargetLanguage
Global $asArrayTabLanguage[1] = [6080]
Global Enum $eCopyright, $eComponents, $eWarranty
Global $asArrayTabAbout[3] = [6096, 6097, 6098]
Global Enum $eCut, $eCopy, $ePaste, $eDelete, $eSelect_All
Global $asArrayContextMenu[5] = [6112, 6113, 6114, 6115, 6116]
Global Enum $eTray_Translate, $eTray_OCR, $eTray_Result, $eExit
Global $asArrayGUI_Tray[4] = [6128, 6129, 6130, 6131]
Global Enum $eAPI_Success, $eAPI_ProKey, $eAPI_NonCorrectKey, $eAPI_Failed, $eNoAPI_Key
Global $asArrayMsg_API[5] = [6160, 6161, 6162, 6163, 6164]
Global Enum $eImageTooLarge
Global $asArrayMsg_Image[1] = [6176]
Global Enum $eOCR_Removed, $eOCR_Saved, $eOCR_Changed, $eTrans_Removed, $eTrans_Saved, $eTrans_Changed, $eOCRFailed_Reg, $eTransFailed_Reg, $eOCRFailed_Unreg, $eTransFailed_Unreg, $eSame
Global $asArrayMsg_Hotkey[11] = [6192, 6193, 6194, 6195, 6196, 6197, 6198, 6199, 6200, 6201, 6202]
Global Enum $e_curl_Download, $e_curl_OCR, $e_curl_Translate
Global $asArrayMsg_curl[3] = [6208, 6209, 6210]
Global Enum $eDownloadError, $eFailedDownload
Global $asArrayMsg_ListLng[2] = [6224, 6225]
Global Enum $eErrorTitle, $eOCRErrorTitle, $eTransErrorTitle, $eParsingError, $eEmptyError, $eTimeoutError, $eUnknownError
Global $asArrayMsg_Error[7] = [6240, 6241, 6242, 6243, 6244, 6245, 6246]
#EndRegion ;---------------------------------------------------------------------------------

Global $sIni_Path = @ScriptDir & "\settings.ini"
Global $iWidth_GUI = 490, $iHeight_GUI = 300
Global $iWinPos_X, $iWinPos_Y, $sFontResult, $iFontSizeResult, $bPutResultToClipboard, $bAlwaysOnTop, $iTrayIconAction, $bIconAnimation, $sOCRAPIKey, $sTranslateAPIKey, $sOCRHotkey, $iOCRHotkeyCode, $sTranslateHotkey, $iTransHotkeyCode, $sTargetLanguage

Global $aisArrayLanguageCodes[191][2] = [[4,"zh-CN"],[1025,"ar"],[1026,"bg"],[1027,"ca"],[1028,"zh-TW"],[1029,"cs"],[1030,"da"],[1031,"de"],[1032,"el"],[1033,"en"],[1034,"es"],[1035,"fi"],[1036,"fr"],[1037,"he"],[1038,"hu"],[1039,"is"],[1040,"it"],[1041,"ja"],[1042,"ko"],[1043,"nl"],[1044,"no"],[1045,"pl"],[1046,"pt"],[1048,"ro"],[1049,"ru"],[1050,"hr"],[1051,"sk"],[1052,"sq"],[1053,"sv"],[1054,"th"],[1055,"tr"],[1056,"ur"],[1057,"id"],[1058,"uk"],[1059,"be"],[1060,"sl"],[1061,"et"],[1062,"lv"],[1063,"lt"],[1064,"tg"],[1065,"fa"],[1066,"vi"],[1067,"hy"],[1068,"az"],[1069,"eu"],[1071,"mk"],[1074,"tn"],[1076,"xh"],[1077,"zu"],[1078,"af"],[1079,"ka"],[1081,"hi"],[1082,"mt"],[1086,"ms"],[1087,"kk"],[1088,"ky"],[1089,"sw"],[1090,"tk"],[1091,"uz"],[1092,"tt"],[1093,"bn"],[1094,"pa"],[1095,"gu"],[1096,"or"],[1097,"ta"],[1098,"te"],[1099,"kn"],[1100,"ml"],[1101,"as"],[1102,"mr"],[1103,"sa"],[1104,"mn"],[1106,"cy"],[1107,"km"],[1108,"lo"],[1110,"gl"],[1111,"gom"],[1115,"si"],[1118,"am"],[1121,"ne"],[1122,"fy"],[1123,"ps"],[1124,"tl"],[1125,"dv"],[1128,"ha"],[1130,"yo"],[1131,"qu"],[1132,"nso"],[1133,"ba"],[1134,"lb"],[1136,"ig"],[1139,"ti"],[1141,"haw"],[1150,"br"],[1152,"ug"],[1153,"mi"],[1154,"oc"],[1155,"co"],[1159,"rw"],[1169,"gd"],[1170,"ku"],[2049,"ar"],[2051,"ca"],[2052,"zh-CN"],[2055,"de"],[2057,"en"],[2058,"es"],[2060,"fr"],[2064,"it"],[2067,"nl"],[2068,"no"],[2070,"pt-PT"],[2077,"sv"],[2080,"ur"],[2092,"az"],[2098,"tn"],[2108,"ga"],[2110,"ms"],[2117,"bn"],[2118,"pa-Arab"],[2121,"ta"],[2128,"mn"],[2137,"sd"],[2151,"ff"],[2155,"qu"],[2163,"ti"],[3073,"ar"],[3076,"zh-HK"],[3079,"de"],[3081,"en"],[3082,"es"],[3084,"fr"],[3098,"sr"],[3179,"qu"],[4097,"ar"],[4100,"zh-CN"],[4103,"de"],[4105,"en"],[4106,"es"],[4108,"fr"],[4122,"hr"],[5121,"ar"],[5124,"zh-TW"],[5127,"de"],[5129,"en"],[5130,"es"],[5132,"fr"],[5146,"bs"],[6145,"ar"],[6153,"en"],[6154,"es"],[6156,"fr"],[7169,"ar"],[7177,"en"],[7178,"es"],[7194,"sr"],[8193,"ar"],[8201,"en"],[8202,"es"],[9217,"ar"],[9225,"en"],[9226,"es"],[10241,"ar"],[10249,"en"],[10250,"es"],[10266,"sr"],[11265,"ar"],[11273,"en"],[11274,"es"],[12289,"ar"],[12297,"en"],[12298,"es"],[12314,"sr"],[13313,"ar"],[13321,"en"],[13322,"es"],[14337,"ar"],[14346,"es"],[15361,"ar"],[15370,"es"],[16385,"ar"],[16393,"en"],[16394,"es"],[17417,"en"],[17418,"es"],[18441,"en"],[18442,"es"],[19466,"es"],[20490,"es"],[21514,"es"],[31748,"zh-TW"]]

_FixAccelHotKeyLayout()
_Check_Exists_INI_File()
_GUI_Window_Position()
$wProcHandle = DllCallbackRegister("_WindowProc", "ptr", "hwnd;uint;wparam;lparam")	;for custom menu

;backup original arrow cursor
$hCursorArrow = _WinAPI_LoadImage(0, $OCR_NORMAL, $IMAGE_CURSOR, 0, 0, BitOR($LR_DEFAULTSIZE, $LR_SHARED))
$hMyCursor = _WinAPI_LoadImage($hInstance, 101, $IMAGE_CURSOR, 0, 0, $LR_DEFAULTSIZE)

#Region ;==================================== TRAY GUI ======================================
$idTrayMenu_Translate = TrayCreateItem(_GetStrRes($asArrayGUI_Tray[$eTray_Translate]))									    ;"Translate"
$idTrayMenu_OCR = TrayCreateItem(_GetStrRes($asArrayGUI_Tray[$eTray_OCR]))												    ;"OCR"
TrayCreateItem("")
$idTrayMenu_Result = TrayCreateItem(_GetStrRes($asArrayGUI_Tray[$eTray_Result]))										    ;"Result"
TrayCreateItem("")
$idTrayMenu_Exit = TrayCreateItem(_GetStrRes($asArrayGUI_Tray[$eExit]))													    ;"Exit"
If $nSystemUsesLightTheme Then
	TraySetIcon(@ScriptFullPath, 202)
Else
	TraySetIcon(@ScriptFullPath, 201)
EndIf

TraySetClick($TRAY_CLICK_SECONDARYUP)
TraySetToolTip (_GetStrRes($asArrayGUI_App[$eAppName]))																	    ;"Capture-OCR-Translate"

_Tray_Icon_Action_Item_Select()
#EndRegion ;----------------------------------------------------------------------

#Region ;====================================== GUI =========================================
$hGUI_App = GUICreate(_GetStrRes($asArrayGUI_App[$eAppName]), $iWidth_GUI, $iHeight_GUI, $iWinPos_X, $iWinPos_Y);"Capture-OCR-Translate"
GUISetFont(10, 400, 0, "Segoe UI")
If $nSystemUsesLightTheme Then
	GUISetIcon(@ScriptFullPath, 202)
Else
	GUISetIcon(@ScriptFullPath, 201)
EndIf

$idTab = GUICtrlCreateTab(10, 10, 470, 280)
;----------------------------------------------------------------------------------TAB RESULT
$idTabResult = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabResult]))												;"Result"
$idEdit_RSLT = GUICtrlCreateEdit("", 30, 48, 430, 228, BitOR($ES_AUTOVSCROLL,$ES_READONLY,$ES_WANTRETURN,$WS_VSCROLL), $WS_EX_STATICEDGE)
GUICtrlSetFont(-1, $iFontSizeResult, 400, 0, $sFontResult)

;-----------------------------------------------------Custom Context Menu For Edit
$hContextMenu_Edit = _GUICtrlMenu_CreatePopup()
If $hContextMenu_Edit <> 0 Then
	_GUICtrlMenu_AddMenuItem($hContextMenu_Edit, _GetStrRes($asArrayContextMenu[$eCopy]), $WM_COPY)                			;"Copy"
	_GUICtrlMenu_AddMenuItem($hContextMenu_Edit, "")
	_GUICtrlMenu_AddMenuItem($hContextMenu_Edit, _GetStrRes($asArrayContextMenu[$eSelect_All]), $idSelectAll)       		;"Select All"
	$hEdit = GUICtrlGetHandle($idEdit_RSLT)
	$wProcOld_Edit = _WinAPI_SetWindowLong($hEdit, $GWL_WNDPROC, DllCallbackGetPtr($wProcHandle))
EndIf
;---------------------------------------------------------------------------------TAB GENERAL
$idTabGeneral = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabGeneral]))												;"General"
$idCheckbox_GNR_Clipboard = GUICtrlCreateCheckbox(_GetStrRes($asArrayTabGeneral[$eClipboard]), 34, 48, 423, 21)				;"Put the result in the clipboard"
If $bPutResultToClipboard Then
	GUICtrlSetState(-1, $GUI_CHECKED)
EndIf

$idCheckbox_GNR_AlwaysOnTop = GUICtrlCreateCheckbox(_GetStrRes($asArrayTabGeneral[$eAlwaysOnTop]), 34, 70, 423, 21)			;"Always on Top"
If $bAlwaysOnTop Then
	GUICtrlSetState(-1, $GUI_CHECKED)
	WinSetOnTop($hGUI_App, "", 1)
EndIf

$idLabel_GNR_TrayIconAction = GUICtrlCreateLabel(_GetStrRes($asArrayTabGeneral[$eTrayIconAction]), 34, 98, 230, 51, $SS_RIGHT)	;"Clicking on the tray icon launches:"
$idCombo_GNR_TrayIconAction = GUICtrlCreateCombo("", 277, 96, 180, 25, BitOR($GUI_SS_DEFAULT_COMBO,$CBS_SIMPLE))
If $iTrayIconAction < 0 Or $iTrayIconAction > (UBound($asArrayTrayIconAction) - 1) Then
	$iTrayIconAction = 0; Default
	IniWrite($sIni_Path, "General", "TrayIconAction", 0)
EndIf
GUICtrlSetData(-1, _Fill_Combobox_TrayIconAction(), _GetStrRes($asArrayTrayIconAction[$iTrayIconAction]))

;---------------------------------------------------------------------------------TAB API KEY
$idTabAPIkey = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabAPIKey]))												;"API Key"
$idLabel_API_Description = GUICtrlCreateLabel(_GetStrRes($asArrayTabAPIKey[$eDescription]), 34, 48, 423, 61)				;"The program requires an API key to work. Get a free API key from the OCR service website and paste the API key in the field below. All you need is a valid email address and 2-3 minutes."
$idButton_API_GetOCRAPIKey = GUICtrlCreateButton(_GetStrRes($asArrayTabAPIKey[$eGetOCRAPIKey]), 34, 116, 205, 31)			;"Get API Key"
$idInput_API_OCRAPIKey = GUICtrlCreateInput("", 252, 119, 205, 25, BitOR($GUI_SS_DEFAULT_INPUT,$ES_RIGHT,$ES_UPPERCASE))
GUICtrlSetLimit($idInput_API_OCRAPIKey, 15)
If Not ($sOCRAPIKey == "") Then
	GUICtrlSetData(-1, $sOCRAPIKey)
Else
	$bNeedAPIKey = True
EndIf
;----------------------------------------------------Custom Context Menu For Input
$hContextMenu_Input = _GUICtrlMenu_CreatePopup()
If $hContextMenu_Input <> 0 Then
	_GUICtrlMenu_AddMenuItem($hContextMenu_Input, _GetStrRes($asArrayContextMenu[$eCut]), $WM_CUT)                  		;"Cut"
	_GUICtrlMenu_AddMenuItem($hContextMenu_Input, _GetStrRes($asArrayContextMenu[$eCopy]), $WM_COPY)                		;"Copy"
	_GUICtrlMenu_AddMenuItem($hContextMenu_Input, _GetStrRes($asArrayContextMenu[$ePaste]), $WM_PASTE)              		;"Paste"
	_GUICtrlMenu_AddMenuItem($hContextMenu_Input, _GetStrRes($asArrayContextMenu[$eDelete]), $WM_CLEAR)             		;"Delete"
	_GUICtrlMenu_AddMenuItem($hContextMenu_Input, "")
	_GUICtrlMenu_AddMenuItem($hContextMenu_Input, _GetStrRes($asArrayContextMenu[$eSelect_All]), $idSelectAll)       		;"Select All"

	$hInput = GUICtrlGetHandle($idInput_API_OCRAPIKey)
	$wProcOld_Input = _WinAPI_SetWindowLong($hInput, $GWL_WNDPROC, DllCallbackGetPtr($wProcHandle))
EndIf
;------------------
$idLabel_API_Hr = GUICtrlCreateLabel("", 11, 238, 466, 1)
GUICtrlSetBkColor(-1, 0xD9D9D9)
$idButton_API_Save = GUICtrlCreateButton(_GetStrRes($asArrayTabAPIKey[$eSave]), 357, 248, 100, 31)							;"Save"
GUICtrlSetState(-1, $GUI_DISABLE)
;---------------------------------------------------------------------------------TAB HOTKEYS
$idTabHotkeys = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabHotkey]))												;"Hotkeys"
$idLabel_HK_HotKeyOCR = GUICtrlCreateLabel(_GetStrRes($asArrayTabHotkey[$eOCRHotkey]), 34, 52, 210, 21, $SS_RIGHT)			;"Hotkey for OCR:"
$hInput_HK_HotKeyOCR = _GUICtrlHotkey_Create($hGUI_App, 257, 50, 200, 25)
_GUICtrlHotkey_SetRules($hInput_HK_HotKeyOCR, BitOR($HKCOMB_S, $HKCOMB_NONE), $HOTKEYF_ALT)
_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)

If $iOCRHotkeyCode <> 0 Then
	_Registration_Hotkey_OCR()
	If @error Then
		$iOCRHotkeyCode = 0
		_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
		IniWrite($sIni_Path, "Hotkey", "OCRHotkeyCode", $iOCRHotkeyCode)
	EndIf
EndIf
$sOCRHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyOCR)

ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
;-------------
$idLabel_HK_HotKeyTrans = GUICtrlCreateLabel(_GetStrRes($asArrayTabHotkey[$eTransHotkey]), 34, 88, 210, 21, $SS_RIGHT)		;"Hotkey for translation:"
$hInput_HK_HotKeyTrans = _GUICtrlHotkey_Create($hGUI_App, 257, 86, 200, 25)
_GUICtrlHotkey_SetRules($hInput_HK_HotKeyTrans, BitOR($HKCOMB_S, $HKCOMB_NONE), $HOTKEYF_ALT)
_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)

If $iTransHotkeyCode <> 0 Then
	_Registration_Hotkey_Translate()
	If @error Then
		$iTransHotkeyCode = 0
		_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
		IniWrite($sIni_Path, "Hotkey", "TranslateHotkeyCode", $iTransHotkeyCode)
	EndIf
EndIf
$sTranslateHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyTrans)

ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
;--------------
$idLabel_HK_Tip = GUICtrlCreateLabel(_GetStrRes($asArrayTabHotkey[$eTipHotkey]), 34, 121, 423, 81)							;"Single characters and Shift-only characters are blocked to prevent software functions from running while typing. Please show a little imagination..."
;--------------
$idLabel_HK_Hr = GUICtrlCreateLabel("", 11, 238, 466, 1)
GUICtrlSetBkColor(-1, 0xD9D9D9)
$idButton_HK_Save = GUICtrlCreateButton(_GetStrRes($asArrayTabAPIKey[$eSave]), 357, 248, 100, 31)                           ;"Save"
GUICtrlSetState(-1, $GUI_DISABLE)
;--------------------------------------------------------------------------------TAB LANGUAGE
$idTabLanguage = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabLanguage]))											;"Language"
$idLabel_LNG_TargetLanguage = GUICtrlCreateLabel(_GetStrRes($asArrayTabLanguage[$eTargetLanguage]), 34, 48, 210, 41, $SS_RIGHT)
$idCombo_LNG_TargetLanguage = GUICtrlCreateCombo("", 257, 54, 200, 25, BitOR($GUI_SS_DEFAULT_COMBO,$CBS_SIMPLE,$CBS_SORT))	;"Target language for translating recognized text:"
Local $sGoogleLangID = _GET_GoogletLngID_from_SystemLngID()
Local $asArrayLanguageCodeAndName = _Download_Translation_Language_List($sGoogleLangID)
Local $sErrorGetList = @error
If $sErrorGetList Then
	Local $sGoogleAPIErrorMessage = ""
	If $sErrorGetList = 42 Then
		$sGoogleAPIErrorMessage = @CRLF & "Google API: "
		If $sGoogleErrorCode <> "" Then
			$sGoogleAPIErrorMessage &=	$sGoogleErrorCode & ": "
		EndIf
		$sGoogleAPIErrorMessage &= $sGoogleError
	EndIf
	MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_ListLng[$eDownloadError]), _GetStrRes($asArrayMsg_ListLng[$eFailedDownload]) & $sGoogleAPIErrorMessage)

	Exit
EndIf
If $sTargetLanguage == "" Then
	$sTargetLanguage = $sGoogleLangID
	IniWrite($sIni_Path, "Language", "TargetLanguage", $sTargetLanguage)
EndIf
GUICtrlSetData(-1, _Fill_Combobox_TargetLanguage(), _Get_Name_TargetLanguage())
;-----------------------------------------------------------------------------------TAB ABOUT
$idTabAbout = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabAbout]))													;"About"
$idIcon_ABT = GUICtrlCreateIcon(@ScriptFullPath, 99, 404, 44, 48, 48)
$idLabel_ABT_AppName = GUICtrlCreateLabel(_GetStrRes($asArrayGUI_App[$eAppName]), 34, 52, 323, 41)							;"Capture-OCR-Translate"
GUICtrlSetFont(-1, 20, 800, 0, "Segoe UI")
GUICtrlSetColor(-1, 0xD9D9D9)
$idLabel_ABT_Version = GUICtrlCreateLabel(FileGetVersion(@ScriptFullPath, "ProductVersion"), 34, 110, 350, 21)
$idLabel_ABT_Copyright = GUICtrlCreateLabel(_GetStrRes($asArrayTabAbout[$eCopyright]), 34, 131, 350, 21)					;"Copyright © NyBumBum 2025–2026"
$idLabel_ABT_Mail = GUICtrlCreateLabel("nybumbum@gmail.com", 34, 152, 350, 21)												;"nybumbum@gmail.com"
GUICtrlSetColor(-1, 0x0078D7)
GUICtrlSetCursor (-1, 0)
$idLabel_ABT_Site = GUICtrlCreateLabel("github.com/nbb1967/capture-ocr-translate", 34, 173, 350, 21)						;"github.com/nbb1967/capture-ocr-translate"
GUICtrlSetColor(-1, 0x0078D7)
GUICtrlSetCursor (-1, 0)
$idLabel_ABT_Components = GUICtrlCreateLabel(_GetStrRes($asArrayTabAbout[$eComponents]), 34, 210, 420, 21)					;"Based on AutoIt, curl, OCR.SPACE and Google Translation."
$idLabel_ABT_Warranty = GUICtrlCreateLabel(_GetStrRes($asArrayTabAbout[$eWarranty]), 34, 230, 422, 41)						;'This free utility is provided "as is" without any warranty...'
$Label1 = GUICtrlCreateLabel("", 11, 98, 466, 1)
GUICtrlSetBkColor(-1, 0xD9D9D9)
GUICtrlCreateTabItem("")

If $bNeedAPIKey Then
	GUISetState(@SW_SHOW, $hGUI_App)
	GUICtrlSetState($idTabAPIkey, $GUI_SHOW)
	$iLastTab = 2
	GUICtrlSetState($idButton_API_GetOCRAPIKey, $GUI_FOCUS)
Else
	GUISetState(@SW_HIDE, $hGUI_App)
	GUICtrlSetState($idTabResult, $GUI_SHOW)
	$iLastTab = 0
EndIf
#EndRegion ;----------------------------------------------------------------------

GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND")		; to track changes in inputs
GUIRegisterMsg($WM_MOVE, "WM_MOVE")				; to track window coordinates when moving

Local $aGUIWindowPosition, $iState
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			;------------------------------------------------ Save Window Position
			; If the window is NOT minimized, save the coordinates
			$iState = WinGetState($hGUI_App)
			If Not BitAND($iState, $WIN_STATE_MINIMIZED) Then
				$aGUIWindowPosition = WinGetPos($hGUI_App)
				If Not @error Then
					IniWrite($sIni_Path, "GUI", "GUIWindowPosition_X", $aGUIWindowPosition[0])
					IniWrite($sIni_Path, "GUI", "GUIWindowPosition_Y", $aGUIWindowPosition[1])
				EndIf
			EndIf
			GUISetState(@SW_HIDE, $hGUI_App)
		Case $idTab	;--------------------------------------------------------- Tab
			$iCurrentTab = GUICtrlRead($idTab)
			If $iCurrentTab <> $iLastTab Then
                If $iLastTab = 2 Then
                    If _WinAPI_GetFocus() = GUICtrlGetHandle($idInput_API_OCRAPIKey) Then
                        GUICtrlSetState($idTab, $GUI_FOCUS)
                    EndIf
                EndIf
                ;------------------

                Switch $iCurrentTab
                    Case 3
                        ControlShow($hGUI_App, "", $hInput_HK_HotKeyOCR)
                        ControlShow($hGUI_App, "", $hInput_HK_HotKeyTrans)

                    Case Else
                        ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
                        ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
                EndSwitch
				$iLastTab = $iCurrentTab
			EndIf

		Case $idCheckbox_GNR_Clipboard	;------------------------------- Clipboard
			If BitAND(GUICtrlRead($idCheckbox_GNR_Clipboard), $GUI_CHECKED) = $GUI_CHECKED Then
				$bPutResultToClipboard = True
			Else
				$bPutResultToClipboard = False
			EndIf
			IniWrite($sIni_Path, "General", "PutResultToClipboard", $bPutResultToClipboard)
		Case $idCheckbox_GNR_AlwaysOnTop	;----------------------- Always on Top
			If BitAND(GUICtrlRead($idCheckbox_GNR_AlwaysOnTop), $GUI_CHECKED) = $GUI_CHECKED Then
				WinSetOnTop($hGUI_App, "", 1)
				$bAlwaysOnTop = True
			Else
				WinSetOnTop($hGUI_App, "", 0)
				$bAlwaysOnTop = False
			EndIf
			IniWrite($sIni_Path, "General", "AlwaysOnTop", $bAlwaysOnTop)
		Case $idCombo_GNR_TrayIconAction
			If Not (GUICtrlRead($idCombo_GNR_TrayIconAction) == _GetStrRes($asArrayTrayIconAction[$iTrayIconAction])) Then
				For $i = 0 To UBound($asArrayTrayIconAction) - 1
					If GUICtrlRead($idCombo_GNR_TrayIconAction) == _GetStrRes($asArrayTrayIconAction[$i]) Then
						$iTrayIconAction = $i
						IniWrite($sIni_Path,  "General", "TrayIconAction", $iTrayIconAction)
						_Tray_Icon_Action_Item_Select()
						ExitLoop
					EndIf
				Next
			EndIf
		Case $idButton_API_Save
			_Check_OCRAPIKey()
		Case $idButton_HK_Save
			_Check_Hotkeys()
		Case $idCombo_LNG_TargetLanguage
			If Not (GUICtrlRead($idCombo_LNG_TargetLanguage) == _Get_Name_TargetLanguage) Then
				For $i = 0 To UBound($asArrayLanguageCodeAndName, $UBOUND_ROWS) - 1
					If GUICtrlRead($idCombo_LNG_TargetLanguage) == $asArrayLanguageCodeAndName[$i][1] Then
						$sTargetLanguage = $asArrayLanguageCodeAndName[$i][0]
						IniWrite($sIni_Path, "Language", "TargetLanguage", $sTargetLanguage)
						ExitLoop
					EndIf
				Next
			EndIf
		Case $idButton_API_GetOCRAPIKey
			ShellExecute("https://ocr.space/ocrapi/freekey")
		Case $idLabel_ABT_Mail
			ShellExecute("mailto:nybumbum@gmail.com?subject=Capture-OCR-Translate")
		Case $idLabel_ABT_Site
			ShellExecute("https://github.com/nbb1967/capture-ocr-translate")
	EndSwitch
	;-------------------------------------------------------------------------Tray
    Local $nTrayMsg = TrayGetMsg()
    Switch $nTrayMsg
		Case $TRAY_EVENT_PRIMARYUP
			_Tray_Icon_Action()
		Case $idTrayMenu_Translate
			_OCR_and_Translate_Tray_Function()
		Case $idTrayMenu_OCR
			_Only_OCR_Tray_Function()
        Case $idTrayMenu_Result
			_Result_Tray_Function()
        Case $idTrayMenu_Exit
            _Exit_Tray_Function()
    EndSwitch
WEnd
;=================================================================================================
;-----------------------------------------------------------CHECKING THAT THE INI FILE EXISTS
Func _Check_Exists_INI_File()
	If Not FileExists($sIni_Path) Then _Create_INI_File()

	_Read_Settings()
EndFunc
;-----------------------------------------------------------------------------CREATE INI FILE
Func _Create_INI_File()
	IniWriteSection($sIni_Path, "GUI", "GUIWindowPosition_X=" & @LF & "GUIWindowPosition_Y=")
	IniWriteSection($sIni_Path, "Result", "FontResult=Segoe UI" & @LF & "FontSizeResult=10")
	IniWriteSection($sIni_Path, "General", "PutResultToClipboard=True" & @LF & "AlwaysOnTop=False" & @LF & "TrayIconAction=0" & @LF & "IconAnimation=True")
	IniWriteSection($sIni_Path, "APIKey", "OCRAPIKey=" & @LF & "TranslateAPIKey=AIzaSyBOti4mM-6x9WDnZIjIeyEU21OpBXqWBgw")
	IniWriteSection($sIni_Path, "Hotkey", "OCRHotkeyCode=" & @LF & "TranslateHotkeyCode=")
	IniWriteSection($sIni_Path, "Language", "TargetLanguage=")
EndFunc
;---------------------------------------------------------READING USER SETTINGS FROM INI FILE
Func _Read_Settings()
	$iWinPos_X = IniRead($sIni_Path, "GUI", "GUIWindowPosition_X", "")
	$iWinPos_Y = IniRead($sIni_Path, "GUI", "GUIWindowPosition_Y", "")
	$sFontResult = IniRead($sIni_Path, "Result", "FontResult", "Segoe UI")
	$iFontSizeResult = IniRead($sIni_Path, "Result", "FontSizeResult", 10)
	$bPutResultToClipboard = (StringLower(IniRead($sIni_Path, "General", "PutResultToClipboard", "True")) == "true")
	$bAlwaysOnTop = (StringLower(IniRead($sIni_Path, "General", "AlwaysOnTop", "False")) == "true")
	$iTrayIconAction = IniRead($sIni_Path, "General", "TrayIconAction", 0)
	$bIconAnimation = (StringLower(IniRead($sIni_Path, "General", "IconAnimation", "True")) ==  "true")
	$sOCRAPIKey = IniRead($sIni_Path, "APIKey", "OCRAPIKey", "")
	$sTranslateAPIKey = IniRead($sIni_Path, "APIKey", "TranslateAPIKey", "AIzaSyBOti4mM-6x9WDnZIjIeyEU21OpBXqWBgw")
	$iOCRHotkeyCode = IniRead($sIni_Path, "Hotkey", "OCRHotkeyCode", "")
	$iTransHotkeyCode = IniRead($sIni_Path, "Hotkey", "TranslateHotkeyCode", "")
	$sTargetLanguage = IniRead($sIni_Path, "Language", "TargetLanguage", "")
EndFunc
;------------------CHECKING THE POSITION OF THE PROGRAM WINDOW. RETURN AN INACCESSIBLE WINDOW
Func _GUI_Window_Position()
    Local $tRECT = _WinAPI_GetWorkArea()
    Local $iLeft = DllStructGetData($tRECT, 'Left')
    Local $iTop = DllStructGetData($tRECT, 'Top')
    Local $iRight = DllStructGetData($tRECT, 'Right')
    Local $iBottom = DllStructGetData($tRECT, 'Bottom')

    If $iWinPos_X <> "" And $iWinPos_Y <> "" Then
        If $iWinPos_X > $iLeft - $iWidth_GUI + 150 And $iWinPos_X < $iRight - 50 And $iWinPos_Y > $iTop - 10 And $iWinPos_Y < $iBottom - 30 Then
            Return
        EndIf
    EndIf

    Local $iHeight_Caption = _WinAPI_GetSystemMetrics($SM_CYCAPTION)
    Local $iSize_Edge = _WinAPI_GetSystemMetrics($SM_CXEDGE)

    $iWinPos_X = $iRight - $iWidth_GUI - (2 * $iSize_Edge)
    $iWinPos_Y = $iBottom - $iHeight_GUI - $iHeight_Caption - (2 * $iSize_Edge)
EndFunc
;----------------------------------------------------------------CUSTOM CONTEXT MENU BY RASYM
Func _WindowProc($hWnd, $Msg, $wParam, $lParam)
	Local $aRet
	Switch $hWnd
		Case $hInput
			Switch $Msg
				Case $WM_CONTEXTMENU
					_WinAPI_SetFocus($hWnd)
					_ActivityContextMenuItem_Input()
					; $wParam in WM_CONTEXTMENU — this is the HWND of the control
					_GUICtrlMenu_TrackPopupMenu($hContextMenu_Input, $hWnd)
					Return 1
				Case $WM_COMMAND
					Local $nID = BitAND($wParam, 0xFFFF) ; Extract the command ID from the message
					Switch $nID
						Case $WM_CUT, $WM_COPY, $WM_PASTE, $WM_CLEAR
							_SendMessage($hWnd, $nID)
						Case $idSelectAll
							_SendMessage($hWnd, $EM_SETSEL, 0, -1)
					EndSwitch
			EndSwitch
			$aRet = DllCall("user32.dll", "int", "CallWindowProc", "ptr", $wProcOld_Input, _
					"hwnd", $hWnd, "uint", $Msg, "wparam", $wParam, "lparam", $lParam)
			Return $aRet[0]
		Case $hEdit
			Switch $Msg
				Case $WM_CONTEXTMENU
					_WinAPI_SetFocus($hWnd)
					_ActivityContextMenuItem_Edit()
					; $wParam in WM_CONTEXTMENU — this is the HWND of the control
					_GUICtrlMenu_TrackPopupMenu($hContextMenu_Edit, $hWnd)
					Return 1
				Case $WM_COMMAND
					Local $nID = BitAND($wParam, 0xFFFF) ; Extract the command ID from the message
					Switch $nID
						Case $WM_COPY
							_SendMessage($hWnd, $nID)
						Case $idSelectAll
							_SendMessage($hWnd, $EM_SETSEL, 0, -1)
					EndSwitch
			EndSwitch
			$aRet = DllCall("user32.dll", "int", "CallWindowProc", "ptr", $wProcOld_Edit, _
					"hwnd", $hWnd, "uint", $Msg, "wparam", $wParam, "lparam", $lParam)
			Return $aRet[0]
	EndSwitch
EndFunc   ;==>_WindowProc

Func WM_MOVE($hWnd, $iMsg, $wParam, $lParam)
	If $hWnd = $hGUI_App Then
		Local $iState = WinGetState($hGUI_App)
		; We save coordinates ONLY if the window is not minimized
		If Not BitAND($iState, $WIN_STATE_MINIMIZED) Then
			Local $aGUIWindowPosition = WinGetPos($hGUI_App)
			If Not @error Then
				IniWrite($sIni_Path, "GUI", "GUIWindowPosition_X", $aGUIWindowPosition[0])
				IniWrite($sIni_Path, "GUI", "GUIWindowPosition_Y", $aGUIWindowPosition[1])
			EndIf
		EndIf
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_MOVE
;--------------------------------------------------------------FILL COMBOBOX TRAY-ICON-ACTION
Func _Fill_Combobox_TrayIconAction()
	Local $sList = ""
	Local $imax = UBound($asArrayTrayIconAction) - 1
	For $i = 0 To $imax
		$sList &= _GetStrRes($asArrayTrayIconAction[$i]) & '|'
	Next
	$sList = StringTrimRight($sList, 1)
	Return $sList

EndFunc   ;==>_Fill_Combobox_TrayIconAction
;-------------------------------------------------------------------------------CHECK API KEY
Func _Check_OCRAPIKey()
	Local $s_tmp_OCRAPIKey
	Local Const $sRegexOCRAPIKey = '(?-i)(K[0-9]{14})'
	Local Const $sRegexOCRAPIKey_Pro = '(?-i)^([A-Z][A-Z0-9]{12})$'
	Local $iOCRAPIKeyLength, $bOCRAPIKeyCorrectness

	$s_tmp_OCRAPIKey = GUICtrlRead($idInput_API_OCRAPIKey)
	If $sOCRAPIKey == $s_tmp_OCRAPIKey Then
		Return
	EndIf

    $bOCRAPIKeyCorrectness = StringRegExp($s_tmp_OCRAPIKey, $sRegexOCRAPIKey)

    If $bOCRAPIKeyCorrectness Then ; Free key
        If IniWrite($sIni_Path, "APIKey", "OCRAPIKey", $s_tmp_OCRAPIKey) Then
            $sOCRAPIKey = $s_tmp_OCRAPIKey
            GUICtrlSetState($idButton_API_Save, $GUI_DISABLE)
            $bNeedAPIKey = False
            MsgBox($MB_OK, "", _GetStrRes($asArrayMsg_API[$eAPI_Success]))														    ;"Free OCR API Key saved successfully."
			GUICtrlSetState($idTab, $GUI_FOCUS)
        Else
            MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Error[$eErrorTitle]), _GetStrRes($asArrayMsg_API[$eAPI_Failed]))			;"Error"	"Failed to save Free OCR API Key."
            GUICtrlSetState($idInput_API_OCRAPIKey, $GUI_FOCUS)
            SetError(14)
        EndIf

    ElseIf StringRegExp($s_tmp_OCRAPIKey, $sRegexOCRAPIKey_Pro) Then ; PRO key
        MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Error[$eErrorTitle]), _GetStrRes($asArrayMsg_API[$eAPI_ProKey]))				;"Error"	"You probably entered a PRO/PRO PDF key. Unfortunately, this program currently only supports Free OCR keys."
        GUICtrlSetState($idInput_API_OCRAPIKey, $GUI_FOCUS)
        SetError(15)

    Else ; Junk
        MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Error[$eErrorTitle]), _GetStrRes($asArrayMsg_API[$eAPI_NonCorrectKey]))		;"Error"	"Please make sure you have entered a valid API key for Free OCR."
        GUICtrlSetState($idInput_API_OCRAPIKey, $GUI_FOCUS)
        SetError(16)
    EndIf

    Return
EndFunc
;----------------------------------------------------------DOWNLOAD TRANSLATION LANGUAGE LIST
Func _Download_Translation_Language_List($sGoogleLangID)
	Local $sCurlCmd = 'curl' & $sSystemProxy & ' -A "Mozilla/5.0" --silent --show-error --ssl-no-revoke -X GET "https://translation.googleapis.com/language/translate/v2/languages?key=' & $sTranslateAPIKey & '&target=' & $sGoogleLangID  & '"'
	Local $iPID = Run($sCurlCmd, "", "", BitOR($STDERR_CHILD, $STDOUT_CHILD))
	;-----------------------------------------------------------------------
	Local $sResultLanguageList = _Curl_WaitAndAnimate($iPID, _GetStrRes($asArrayMsg_ListLng[$eDownloadError]), _GetStrRes($asArrayMsg_Error[$eTimeoutError]), _GetStrRes($asArrayMsg_curl[$e_curl_Download]), 12) ; Offset = 12, will give @error 14 (12+2) or 15 (12+3)
	If @error Then
		; Throwing errors 14 or 15 higher
		Return SetError(@error, 0, 0)
	EndIf
	;------------------------------------------------------------------------------Parsing List
	Local $sJSON_LangugeName = _WinAPI_MultiByteToWideChar($sResultLanguageList, 65001, 0, True)

    $sGoogleError = ""
    $sGoogleErrorCode = ""
	Local $mJson = _JSON_Parse($sJSON_LangugeName)
	If @error Then Return SetError(41, 0, "")

	Local $aLanguages = _JSON_Get($mJson, "data.languages")
    If @error Then
        _Check_Google_Error($mJson)
        If @error Then Return SetError(@error, 0, "") ; Error forwarding
        Return SetError(42, 0, "")
    EndIf

	If Not IsArray($aLanguages) Then Return SetError(44, 0, "")

	Local $iRows = UBound($aLanguages)
	Local $asArrayLanguageCodeAndName[$iRows][2]

	For $i = 0 To $iRows - 1
		$asArrayLanguageCodeAndName[$i][0] = _JSON_Get($aLanguages[$i], "language")
		$asArrayLanguageCodeAndName[$i][1] = _JSON_Get($aLanguages[$i], "name")
	Next

	Return $asArrayLanguageCodeAndName
EndFunc ;==>_Download_Translation_Language_List
;----------------------------------------------GET GOOGLE LANGUAGE ID FROM SYSTEM LANGUAGE ID
Func _GET_GoogletLngID_from_SystemLngID()
	Local $iMUI_ID = _WinAPI_GetSystemDefaultUILanguage()
	Local $sGoogleLangID = "en"
	Local $imax = UBound($aisArrayLanguageCodes, $UBOUND_ROWS) - 1

	For $i = 0 To $imax
		if $aisArrayLanguageCodes[$i][0] = $iMUI_ID Then
			$sGoogleLangID = $aisArrayLanguageCodes[$i][1]
			ExitLoop
		EndIf

	Next
	Return $sGoogleLangID

EndFunc ;==>_GET_GoogletLngID_from_SystemLngID
;---------------------------------------------------------------FILL COMBOBOX TARGET LANGUAGE
Func _Fill_Combobox_TargetLanguage()
	Local $sList = ""
	Local $imax = UBound($asArrayLanguageCodeAndName, $UBOUND_ROWS) - 1
	For $i = 0 To $imax
		$sList &= $asArrayLanguageCodeAndName[$i][1] & '|'
	Next
	$sList = StringTrimRight($sList, 1)
	Return $sList

EndFunc   ;==>_Fill_Combobox_TargetLanguage
;--------------------------------------------------------------------GET NAME TARGET LANGUAGE
Func _Get_Name_TargetLanguage()
	Local $sName_TargetLanguage = ""
	Local $imax = UBound($asArrayLanguageCodeAndName, $UBOUND_ROWS) - 1
	For $i = 0 To $imax
		If $sTargetLanguage == $asArrayLanguageCodeAndName[$i][0] Then
			$sName_TargetLanguage = $asArrayLanguageCodeAndName[$i][1]
			ExitLoop
		EndIf
	Next
	Return $sName_TargetLanguage
EndFunc	;==>_Get_Name_TargetLanguage
;----------------------------------------------------START OCR WITH HOTKEY (BASED ON CREATOR)
Func _Only_OCR_Hotkey_Function()
	Local $sHotkey = @HotKeyPressed
	HotKeySet($sHotkey)
	While _IsPressed("10", $hDLL) Or _IsPressed("11", $hDLL) Or _IsPressed("12", $hDLL)
		Sleep(10)
	WEnd

	_Only_OCR_Tray_Function()
	HotKeySet($sHotkey, "_Only_OCR_Hotkey_Function")
EndFunc	;==>_Only_OCR_Hotkey_Function
;------------------------------------START OCR AND TRANSLATION WITH HOTKEY (BASED ON CREATOR)
Func _OCR_and_Translate_Hotkey_Function()
	Local $sHotkey = @HotKeyPressed
	HotKeySet($sHotkey)
	While _IsPressed("10", $hDLL) Or _IsPressed("11", $hDLL) Or _IsPressed("12", $hDLL)
		Sleep(10)
	WEnd

	_OCR_and_Translate_Tray_Function()
	HotKeySet($sHotkey, "_OCR_and_Translate_Hotkey_Function")
EndFunc	;==>_OCR_and_Translate_Hotkey_Function
;----------------------------------------------------------FIX ACCEL HOTKEY LAYOUT BY CREATOR
Func _FixAccelHotKeyLayout()
	Static $iKbrdLayout, $aKbrdLayouts

	If Execute('@exitMethod') <> '' Then
		Local $iUnLoad = 1

		For $i = 1 To UBound($aKbrdLayouts) - 1
			If Hex($iKbrdLayout) = Hex('0x' & StringRight($aKbrdLayouts[$i], 4)) Then
				$iUnLoad = 0
				ExitLoop
			EndIf
		Next

		If $iUnLoad Then
			_WinAPI_UnloadKeyboardLayout($iKbrdLayout)
		EndIf

		Return
	EndIf

	$iKbrdLayout = 0x0409
	$aKbrdLayouts = _WinAPI_GetKeyboardLayoutList()
	_WinAPI_LoadKeyboardLayout($iKbrdLayout, $KLF_ACTIVATE)

	OnAutoItExitRegister('_FixAccelHotKeyLayout')
EndFunc   ;==>_FixAccelHotKeyLayout
;---------------------------------------------------------------------------CHECK ALL HOTKEYS
Func _Check_Hotkeys()
	_FixAccelHotKeyLayout()
	Local $i_tmp_OCRHotkeyCode = _GUICtrlHotkey_GetHotkeyCode($hInput_HK_HotKeyOCR)
	Local $i_tmp_TransHotkeyCode = _GUICtrlHotkey_GetHotkeyCode($hInput_HK_HotKeyTrans)

	If $i_tmp_OCRHotkeyCode = $i_tmp_TransHotkeyCode And $i_tmp_OCRHotkeyCode <> 0 Then
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Error[$eErrorTitle]), _GetStrRes($asArrayMsg_Hotkey[$eSame]))		;"Error"	"The hotkeys for two functions cannot be the same."
		_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
		_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
		GUICtrlSetState($idButton_HK_Save, $GUI_DISABLE)
		SetError(20)
		Return
	EndIf

	_Check_OCR_HK($i_tmp_OCRHotkeyCode)
	_Check_Trans_HK($i_tmp_TransHotkeyCode)

EndFunc   ;==>_Check_Hotkeys
;----------------------------------------------------------------------------CHECK OCR HOTKEY
Func _Check_OCR_HK($i_tmp_OCRHotkeyCode)
	If $i_tmp_OCRHotkeyCode = $iOCRHotkeyCode Then										;nothing happened
		Return
	Else
		If $i_tmp_OCRHotkeyCode = 0 And $iOCRHotkeyCode <> 0 Then						;remove hotkey
			_Unregistration_Hotkey_OCR()
			if @error Then
				_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
				Return
			Else
				IniWrite($sIni_Path, "Hotkey", "OCRHotkeyCode", $i_tmp_OCRHotkeyCode)
				MsgBox($MB_OK, "", _GetStrRes($asArrayMsg_Hotkey[$eOCR_Removed]))								;"Hotkey for OCR successfully removed."
				$iOCRHotkeyCode = $i_tmp_OCRHotkeyCode ; 0
				_Checking_Changes_Hotkeys()
				$sOCRHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyOCR) ; ""
				Return
			EndIf
		ElseIf $i_tmp_OCRHotkeyCode <> 0 And $iOCRHotkeyCode = 0 Then					;create hotkey
			_Registration_Hotkey_OCR()
			If @error Then
				_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
				Return
			Else
				IniWrite($sIni_Path, "Hotkey", "OCRHotkeyCode", $i_tmp_OCRHotkeyCode)
				MsgBox($MB_OK, "", _GetStrRes($asArrayMsg_Hotkey[$eOCR_Saved]))									;"Hotkey for OCR successfully saved."
				$iOCRHotkeyCode = $i_tmp_OCRHotkeyCode
				_Checking_Changes_Hotkeys()
				$sOCRHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyOCR)
				Return
			EndIf
		ElseIf $i_tmp_OCRHotkeyCode <> 0 And $iOCRHotkeyCode <> 0 Then					;change hotkey
			_Registration_Hotkey_OCR()
			If @error Then
				_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
				Return
			Else
				_Unregistration_Hotkey_OCR()
				IniWrite($sIni_Path, "Hotkey", "OCRHotkeyCode", $i_tmp_OCRHotkeyCode)							;"Hotkey for OCR successfully changed."
				MsgBox($MB_OK, "", _GetStrRes($asArrayMsg_Hotkey[$eOCR_Changed]))
				$iOCRHotkeyCode = $i_tmp_OCRHotkeyCode
				_Checking_Changes_Hotkeys()
				$sOCRHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyOCR)
			EndIf
		EndIf
	EndIf
EndFunc   ;==>_Check_OCR_HK
;----------------------------------------------------------------------CHECK TRANSLATE HOTKEY
Func _Check_Trans_HK($i_tmp_TransHotkeyCode)
	If $i_tmp_TransHotkeyCode = $iTransHotkeyCode Then									;nothing happened
		Return
	Else
		If $i_tmp_TransHotkeyCode = 0 And $iTransHotkeyCode <> 0 Then					;remove hotkey
			_Unregistration_Hotkey_Translate()
			If @error Then
				_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
				Return
			Else
				IniWrite($sIni_Path, "Hotkey", "TranslateHotkeyCode", $i_tmp_TransHotkeyCode)
				MsgBox($MB_OK, "", _GetStrRes($asArrayMsg_Hotkey[$eTrans_Removed]))								;"Hotkey for translation successfully removed."
				$iTransHotkeyCode = $i_tmp_TransHotkeyCode ; 0
				_Checking_Changes_Hotkeys()
				$sTranslateHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyTrans) ; ""
				Return
			EndIf
		ElseIf $i_tmp_TransHotkeyCode <> 0 And $iTransHotkeyCode = 0 Then				;create hotkey
			_Registration_Hotkey_Translate()
			If @error Then
				_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
				Return
			Else
				IniWrite($sIni_Path, "Hotkey", "TranslateHotkeyCode", $i_tmp_TransHotkeyCode)
				MsgBox($MB_OK, "", _GetStrRes($asArrayMsg_Hotkey[$eTrans_Saved]))								;"Hotkey for translation successfully saved."
				$iTransHotkeyCode = $i_tmp_TransHotkeyCode
				_Checking_Changes_Hotkeys()
				$sTranslateHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyTrans)
				Return
			EndIf
		ElseIf $i_tmp_TransHotkeyCode <> 0 And $iTransHotkeyCode <> 0 Then				;change hotkey
			_Registration_Hotkey_Translate()
			If @error Then
				_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
				Return
			Else
				_Unregistration_Hotkey_Translate()
				IniWrite($sIni_Path, "Hotkey", "TranslateHotkeyCode", $i_tmp_TransHotkeyCode)
				MsgBox($MB_OK, "", _GetStrRes($asArrayMsg_Hotkey[$eTrans_Changed]))								;"Hotkey for translation successfully changed."
				$iTransHotkeyCode = $i_tmp_TransHotkeyCode
				_Checking_Changes_Hotkeys()
				$sTranslateHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyTrans)
			EndIf
		EndIF
	EndIf
EndFunc   ;==>_Check_Trans_HK
;---------------------------------------------------------------------REGISTRATION OCR HOTKEY
Func _Registration_Hotkey_OCR()
	Local $s_TMP_OCRHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyOCR)
	Local $bSuccess = HotKeySet(StringLower($s_TMP_OCRHotkey), "_Only_OCR_Hotkey_Function")
	If Not $bSuccess Then
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Error[$eErrorTitle]), _GetStrRes($asArrayMsg_Hotkey[$eOCRFailed_Reg]))		;"Error"	"Hotkey for OCR could not be registered in Windows (they may be in use by another program). Try using a different hotkey for OCR."
		SetError(17)
	EndIf
EndFunc   ;==>_Registration_Hotkey_OCR
;---------------------------------------------------------------REGISTRATION TRANSLATE HOTKEY
Func _Registration_Hotkey_Translate()
	Local $s_TMP_TransHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyTrans)
	Local $bSuccess = HotKeySet(StringLower($s_TMP_TransHotkey), "_OCR_and_Translate_Hotkey_Function")
	If Not $bSuccess Then
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Error[$eErrorTitle]), _GetStrRes($asArrayMsg_Hotkey[$eTransFailed_Reg]))	;"Error" 	"Hotkey for translation could not be registered in Windows (they may be in use by another program). Try using a different hotkey for translation."
		SetError(18)
	EndIf
EndFunc   ;==>_Registration_Hotkey_Translate
;-------------------------------------------------------------------UNREGISTRATION OCR HOTKEY
Func _Unregistration_Hotkey_OCR()
	Local $bSuccess = HotKeySet(StringLower($sOCRHotkey))
	If Not $bSuccess Then
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Error[$eErrorTitle]), _GetStrRes($asArrayMsg_Hotkey[$eOCRFailed_Unreg]))	;"Error"	"Failed to unregister hotkey for OCR."
		SetError(21)
	EndIf
EndFunc   ;==>_Unregistration_Hotkey_OCR
;-------------------------------------------------------------UNREGISTRATION TRANSLATE HOTKEY
Func _Unregistration_Hotkey_Translate()
	Local $bSuccess = HotKeySet(StringLower($sTranslateHotkey))
	If Not $bSuccess Then
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Error[$eErrorTitle]), _GetStrRes($asArrayMsg_Hotkey[$eTransFailed_Unreg]))	;"Error"	"Failed to unregister hotkey for translation."
		SetError(22)
	EndIf
EndFunc   ;==>_Unregistration_Hotkey_Translate
;---------------------------------------------------------ACTION BY CLICKING ON THE TRAY ICON
Func _Tray_Icon_Action()
	Switch $iTrayIconAction
		Case 0
			_OCR_and_Translate_Tray_Function()
		Case 1
			_Only_OCR_Tray_Function()
		Case 2
			_Result_Tray_Function()
	EndSwitch
EndFunc   ;==>_Tray_Icon_Action
;---------------------------------SELECTING ITEM MENU FOR ACTION BY CLICKING ON THE TRAY ICON
Func _Tray_Icon_Action_Item_Select()
	Switch $iTrayIconAction
		Case 0
			TrayItemSetState ($idTrayMenu_Translate, $TRAY_DEFAULT)
		Case 1
			TrayItemSetState ($idTrayMenu_OCR, $TRAY_DEFAULT)
		Case 2
			TrayItemSetState ($idTrayMenu_Result, $TRAY_DEFAULT)
	EndSwitch
EndFunc   ;==>_Tray_Icon_Action_Item_Select
;----------------------------------------------------------------------TRANSLATE IN TRAY MENU
Func _OCR_and_Translate_Tray_Function()
	If $bIsProcessing Then Return
    $bIsProcessing = True

	Local $aiPos = _ScreenRegion_GetRect(0xFF0078D7, 0x460066CC, 1)
	If @error Then
		_Clearing_Task_Queue()
		Return
	EndIf
	Local $hCapturedBitmap = _ScreenCapture_Capture("", $aiPos[0], $aiPos[1], $aiPos[2], $aiPos[3], False)
	Local $sWindowsFullPathToImageFile = _SaveCapturedBitmap($hCapturedBitmap)
	If @error Then
		_Clearing_Task_Queue()
		Return
	EndIf

	Local $sOCR_OutText = _SendTo_OCR_Service($sWindowsFullPathToImageFile)
	If @error Then
		_Clearing_Task_Queue()
		Return
	EndIf

	Local $sTranslateOutText = _SendTo_Translate_Service($sOCR_OutText)
	If @error Then
		_Clearing_Task_Queue()
		Return
	EndIf

	If $bPutResultToClipboard Then
		ClipPut($sTranslateOutText)
	EndIf

	Local $iState = WinGetState($hGUI_App)
	If BitAND($iState, $WIN_STATE_MINIMIZED) Then
		WinSetState($hGUI_App, "", @SW_RESTORE)
	EndIf
	GUICtrlSetData($idEdit_RSLT, $sTranslateOutText)
	GUISetState(@SW_SHOW, $hGUI_App)
	If $iLastTab = 3 Then
		ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
		ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
	EndIf
	GUICtrlSetState($idTabResult, $GUI_SHOW)
	$iLastTab = 0
	WinActivate($hGUI_App)

	_Clearing_Task_Queue()
	;-------------------------->>>	set focus?
EndFunc	;==>_OCR_and_Translate_Tray_Function
;----------------------------------------------------------------------------OCR IN TRAY MENU
Func _Only_OCR_Tray_Function()
	If $bIsProcessing Then Return
    $bIsProcessing = True

	Local $aiPos = _ScreenRegion_GetRect(0xFF0078D7, 0x460066CC, 1)
	If @error Then
		_Clearing_Task_Queue()
		Return
	EndIf
	Local $hCapturedBitmap = _ScreenCapture_Capture("", $aiPos[0], $aiPos[1], $aiPos[2], $aiPos[3], False)
	Local $sWindowsFullPathToImageFile = _SaveCapturedBitmap($hCapturedBitmap)
	If @error Then
		_Clearing_Task_Queue()
		Return
	EndIf

	Local $sOCR_OutText = _SendTo_OCR_Service($sWindowsFullPathToImageFile)
	If @error Then
		_Clearing_Task_Queue()
		Return
	EndIf

	If $bPutResultToClipboard Then
		ClipPut($sOCR_OutText)
	EndIf

	Local $iState = WinGetState($hGUI_App)
	If BitAND($iState, $WIN_STATE_MINIMIZED) Then
		WinSetState($hGUI_App, "", @SW_RESTORE)
	EndIf
	GUICtrlSetData($idEdit_RSLT, $sOCR_OutText)
	GUISetState(@SW_SHOW, $hGUI_App)
	If $iLastTab = 3 Then
		ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
		ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
	EndIf
	GUICtrlSetState($idTabResult, $GUI_SHOW)
	$iLastTab = 0
	WinActivate($hGUI_App)

	_Clearing_Task_Queue()
	;-------------------------->>>	set focus?
EndFunc	;==>_Only_OCR_Tray_Function
;-------------------------------------------------------------------------RESULT IN TRAY MENU
Func _Result_Tray_Function()
	Local $iState = WinGetState($hGUI_App)

	If BitAND($iState, $WIN_STATE_MINIMIZED) Then
		GUISetState(@SW_RESTORE, $hGUI_App)
	EndIf

	If $iLastTab = 3 Then
		ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
		ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
	EndIf

	$iLastTab = 0
	GUICtrlSetState($idTabResult, $GUI_SHOW) ; <--- Возвращена на место

	GUISetState(@SW_SHOW, $hGUI_App)
	WinActivate($hGUI_App)
EndFunc
;------------------------------------------------------------------------------------CLEANING
Func _Cleaning()
	;Delete temporary files if any
	Local $bSuccess = FileExists($sPng_Path)
	if $bSuccess Then
		FileDelete($sPng_Path)
	EndIf

	$bSuccess = FileExists($sJpg_Path)
	if $bSuccess Then
		FileDelete($sJpg_Path)
	EndIf

	DllClose($hDLL)

	GUIDelete($hGUI_App)

	;Delete Custom Menu
	DllCallbackFree($wProcHandle)

	;Delete Hotkey Inputs
	_GUICtrlHotkey_Delete($hInput_HK_HotKeyOCR)
	_GUICtrlHotkey_Delete($hInput_HK_HotKeyTrans)

	;Delete Cursors
	_WinAPI_DestroyCursor($hCursorArrow)
	_WinAPI_DestroyCursor($hMyCursor)
EndFunc ;==>_Cleaning
;----------------------------------------------------------------------------------------EXIT
Func _Exit_Tray_Function()
	_Cleaning()
	Exit
EndFunc	;==>_Exit_Tray_Function
;--------------------------------------------------------------------GETTING THE CAPTURE AREA
Func _ScreenRegion_GetRect($nEdge, $nFill, $iSize)

	Local $hCopyMyCursor = _WinAPI_CopyCursor($hMyCursor)
	Local $hCopyCursorArrow = _WinAPI_CopyCursor($hCursorArrow)
	_WinAPI_SetSystemCursor($hCopyMyCursor, $OCR_NORMAL);+destroy object cursor

	;------------------------------------------------------------------------
	_GDIPlus_Startup()

	Local $hCursorGUI = GUICreate("Screen Capture", @DesktopWidth, @DesktopHeight, 0, 0, $WS_POPUP, $WS_EX_TOPMOST)
	WinSetTrans($hCursorGUI, "", 1)
	GUISetState()

	Local $hCapture = GUICreate("", 0, 0, 0, 0, $WS_POPUPWINDOW, BitOR($WS_EX_LAYERED, $WS_EX_TOPMOST, $WS_EX_TOOLWINDOW), $hCursorGUI)
	GUISetState()

	While Sleep(10)
		If _IsPressed("01", $hDLL) Then
			ExitLoop
		ElseIf _IsPressed("1B", $hDLL) Or _IsPressed("02", $hDLL) Then
			While _IsPressed("1B", $hDLL) Or _IsPressed("02", $hDLL)     ; updown ESC or Right Mouse Button
				Sleep(10)
			WEnd
			SetError(30)
			ExitLoop
		EndIf
	WEnd

	If @error Then
		;restore original arrow cursor
		_WinAPI_SetSystemCursor($hCopyCursorArrow, $OCR_NORMAL);+destroy object cursor
		GUIDelete($hCursorGUI)
		GUIDelete($hCapture)
		_GDIPlus_Shutdown()
		SetError(31)
		Return
	EndIf

	Local $iX1 = MouseGetPos(0), $iY1 = MouseGetPos(1)
	Local $iX2, $iY2, $iPosX, $iPosY, $iWidth, $iHeight

	While _IsPressed("01", $hDLL) And Sleep(10)
		$iX2 = MouseGetPos(0)
		$iY2 = MouseGetPos(1)

		Local $aCorrect = _Correct_Capture_Direction($iX1, $iY1, $iX2, $iY2)
		$iPosX = $aCorrect[0]
		$iPosY = $aCorrect[1]
		$iWidth = $aCorrect[2]
		$iHeight = $aCorrect[3]

		Rect($hCapture, $iPosX, $iPosY, $iWidth, $iHeight, $nEdge, $nFill, $iSize)
	WEnd
	Local $aReturn = [$iPosX, $iPosY, $iPosX + $iWidth, $iPosY + $iHeight]
	;---------------------------------------------------------------------------
	;restore original arrow cursor
	_WinAPI_SetSystemCursor($hCopyCursorArrow, $OCR_NORMAL);+destroy object cursor

	GUIDelete($hCursorGUI)
	GUIDelete($hCapture)
	_GDIPlus_Shutdown()

	Return $aReturn
EndFunc   ;==>_ScreenRegion_GetRect
;----------------------------------------------------------------------BLUE SELECTION BY NINE
Func Rect($hWnd, $iPosX, $iPosY, $iWidth, $iHeight, $nEdge, $nFill, $iSize)
  Local $hImage = _GDIPlus_BitmapCreateFromScan0($iWidth, $iHeight)
  Local $hGraphic = _GDIPlus_ImageGetGraphicsContext($hImage)

  Local $hPen = _GDIPlus_PenCreate($nEdge, $iSize)
  Local $hBrush = _GDIPlus_BrushCreateSolid($nFill)

  _GDIPlus_GraphicsFillRect($hGraphic, 0, 0, $iWidth, $iHeight, $hBrush)
  _GDIPlus_GraphicsDrawRect($hGraphic, 0, 0, $iWidth - $iSize, $iHeight - $iSize, $hPen)

  _GDIPlus_PenDispose($hPen)
  _GDIPlus_BrushDispose($hBrush)

  Local $hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hImage)
  _GDIPlus_ImageDispose($hImage)
  _GDIPlus_GraphicsDispose($hGraphic)

  _WinAPI_SetWindowPos($hWnd, 0, $iPosX, $iPosY, $iWidth, $iHeight, $SWP_NOZORDER)
  _WinAPI_UpdateLayeredWindowEx($hWnd, -1, -1, $hBitmap, 255, True)
EndFunc   ;==>Rect
;----------------------------------------------------------------CORRECTING CAPTURE DIRECTION
Func _Correct_Capture_Direction($iX1, $iY1, $iX2, $iY2)
	Local $iPosX, $iPosY, $iWidth, $iHeight
	If  $iX2 < $iX1 Then
        $iPosX = $iX2
        $iWidth = $iX1 - $iX2
    Else
        $iPosX = $iX1
        $iWidth = $iX2 - $iX1
    EndIf

    If $iY2 < $iY1 Then
        $iPosY = $iY2
        $iHeight = $iY1 - $iY2
    Else
        $iPosY = $iY1
        $iHeight = $iY2 - $iY1
    EndIf
	Local $aReturn = [$iPosX, $iPosY, $iWidth, $iHeight]
	Return $aReturn
EndFunc ;==>_Correct_Capture_Direction
;----------------------------------------------------------------------SAVING CAPTURED BITMAP
Func _SaveCapturedBitmap($hCapturedBitmap)
	_ScreenCapture_SaveImage($sPng_Path, $hCapturedBitmap, False)
	Local $iFileSize = FileGetSize($sPng_Path)
	If $iFileSize > 1048576 Then
		Const $iQuality = 85
		_ScreenCapture_SetJPGQuality($iQuality)
		_ScreenCapture_SaveImage($sJpg_Path, $hCapturedBitmap)
		$iFileSize = FileGetSize($sJpg_Path)
		If $iFileSize > 1048576 Then
			MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_Error[$eErrorTitle]), _GetStrRes($asArrayMsg_Image[$eImageTooLarge]))				;"Error", "Image is too large, try reducing capture area."
			SetError(1)
			Return
		Else
			$sImageForOCR_Path = $sJpg_Path
		EndIf
	Else
		$sImageForOCR_Path = $sPng_Path
		_WinAPI_DeleteObject($hCapturedBitmap)
	EndIf
	$sWindowsFullPathToImageFile = _PathFull($sImageForOCR_Path)
	Return $sWindowsFullPathToImageFile
EndFunc ;==>_SaveCapturedBitmap
;----------------------------------------------------------------------URI ENCODE BY PROG@NDY
Func _URIEncode($sData)
    Local $aData = StringSplit(BinaryToString(StringToBinary($sData,4),1),"")
    Local $nChar
    $sData=""
    For $i = 1 To $aData[0]
        $nChar = Asc($aData[$i])
        Switch $nChar
            Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
                $sData &= $aData[$i]
            Case 32
                $sData &= "%20"		;for Google Translate
            Case Else
                $sData &= "%" & Hex($nChar,2)
        EndSwitch
    Next
    Return $sData
EndFunc ;==>_URIEncode
;----------------------------------------------------------------------SENDING TO OCR SERVICE
Func _SendTo_OCR_Service($sWindowsFullPathToImageFile)
	If $bNeedAPIKey Then
		MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_Error[$eErrorTitle]), _GetStrRes($asArrayMsg_API[$eNoAPI_Key]))		; "Error"	"The program cannot work without an API key."
		Local $iState = WinGetState($hGUI_App)
		If BitAND($iState, $WIN_STATE_MINIMIZED) Then
			WinSetState($hGUI_App, "", @SW_RESTORE)
		EndIf
		GUISetState(@SW_SHOW, $hGUI_App)
		If $iLastTab = 3 Then
			ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
			ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
		EndIf
		GUICtrlSetState($idTabAPIkey, $GUI_SHOW)
		$iLastTab = 2
		GUICtrlSetState($idButton_API_GetOCRAPIKey, $GUI_FOCUS)
		SetError(32)
		Return
	EndIf
	;-----------------------------------------------Sending the image to OCR
	$sUnixFullPathToImageFile = StringReplace($sWindowsFullPathToImageFile, "\", "/")
	$sSystemProxy = _GetSystemProxyForCurl()
	$sCurlCmd = 'curl' & $sSystemProxy & ' -A "Mozilla/5.0" --silent --show-error -H "apikey:' & $sOCRAPIKey & '" --ssl-no-revoke --form "file=@' & $sUnixFullPathToImageFile & '" --form "OCREngine=2" --form "language=auto" --form "scale=true" --form "detectOrientation=true" https://api.ocr.space/Parse/Image'
	$iPID = Run($sCurlCmd, "", "", BitOR($STDERR_CHILD, $STDOUT_CHILD))
	;----------------------------------------------------------------------
	Local $sResultOCR = _Curl_WaitAndAnimate($iPID, _GetStrRes($asArrayMsg_Error[$eOCRErrorTitle]), _GetStrRes($asArrayMsg_Error[$eTimeoutError]), _GetStrRes($asArrayMsg_curl[$e_curl_OCR]), 0) ; Offset = 0, will give @error 2 or 3
	If @error Then
        ; Throwing errors 2 or 3 higher
        Return SetError(@error, 0, 0)
    EndIf
	;---------------------------------------------------OCR Parsing & Error
	Local $mJson = _JSON_Parse($sResultOCR)
    If @error Then
		MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_Error[$eOCRErrorTitle]), _GetStrRes($asArrayMsg_Error[$eParsingError]))				;"OCR error"	"Unable to parse the server response."
		Return SetError(4, 0, "")
	EndIf

	$sOCRText = _JSON_Get($mJson, "ParsedResults[0].ParsedText")
	If @error Then
		Local $sOCRErrorMessage = _JSON_Get($mJson, "ParsedResults[0].ErrorMessage")
		If @error Or $sOCRErrorMessage == "" Then
			MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_Error[$eOCRErrorTitle]), _GetStrRes($asArrayMsg_Error[$eUnknownError]))			;"OCR error"	"An unknown error occurred."
			Return SetError(5, 0, "")
		EndIf
		MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_Error[$eOCRErrorTitle]), $sOCRErrorMessage)											;"OCR error"	OCR API Error Message
		Return SetError(6, 0, "")
	EndIf
	If $sOCRText == "" Then
		MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_Error[$eOCRErrorTitle]), _GetStrRes($asArrayMsg_Error[$eEmptyError]))				    ;"OCR error"	"The query result is empty."
		Return SetError(7, 0, "")
	EndIf
	;--------------------------------------------------------------------
	Local $sDecodeParsedTextTo_UTF8 = _WinAPI_MultiByteToWideChar($sOCRText, 65001, 0, True)
	Const $sRegexLineBreak = '([.?!:][\x22\xBB\x{201D}\x{2019}\]\)]?)\h?\n'	; '(?<=[.?!:;])\n'
	Local $sOCR_OutText =  StringRegExpReplace($sDecodeParsedTextTo_UTF8, $sRegexLineBreak, '$1##BREAK##')
	$sOCR_OutText = StringReplace($sOCR_OutText, '-' & @LF, '-')
	$sOCR_OutText = StringReplace($sOCR_OutText, @LF, ' ')
	$sOCR_OutText = StringReplace($sOCR_OutText, '\"', '"')
	$sOCR_OutText = StringReplace($sOCR_OutText, '\\', '\')
	$sOCR_OutText = StringReplace($sOCR_OutText, '\/', '/')   ;\b,\r,\f,\t
	$sOCR_OutText = StringReplace($sOCR_OutText, '##BREAK##', @CRLF)

	Return $sOCR_OutText
EndFunc ;==>_SendTo_OCR_Service
;----------------------------------------------------------------SENDING TO TRANSLATE SERVICE
Func _SendTo_Translate_Service($sOCR_OutText)

	Local $sURI_TextForTranslate = _URIEncode($sOCR_OutText)
	Local $sCurlCmd = 'curl' & $sSystemProxy & ' -A "Mozilla/5.0" --silent --show-error --ssl-no-revoke -X GET "https://www.googleapis.com/language/translate/v2?key=' & $sTranslateAPIKey & '&target=' & $sTargetLanguage & '&format=text&q=' & $sURI_TextForTranslate & '"'
	Local $iPID = Run($sCurlCmd, "", "", BitOR($STDERR_CHILD, $STDOUT_CHILD))
	;------------------------------------------------------------------------------curl
	$sResultTrans = _Curl_WaitAndAnimate($iPID, _GetStrRes($asArrayMsg_Error[$eTransErrorTitle]), _GetStrRes($asArrayMsg_Error[$eTimeoutError]), _GetStrRes($asArrayMsg_curl[$e_curl_Translate]), 6) ; Offset = 6, will give @error 8 (6+2) or 9 (6+3)
    If @error Then
        ; Throwing errors 8 or 9 higher
        Return SetError(@error, 0, 0)
    EndIf
	;------------------------------------------------------Translation Parsing & Errors
	$sGoogleError = ""
    $sGoogleErrorCode = ""

    Local $mJson = _JSON_Parse($sResultTrans)
    If @error Then
		MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_Error[$eTransErrorTitle]), _GetStrRes($asArrayMsg_Error[$eParsingError]))				;"Translation error"	"Unable to parse the server response."
		Return SetError(13, 0, "")
	EndIf

    ; Пробуем достать перевод
    Local $sTranslatedText = _JSON_Get($mJson, "data.translations[0].translatedText")

    If @error Then
        _Check_Google_Error($mJson)
        If @error Then
			MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_Error[$eTransErrorTitle]), _GetStrRes($asArrayMsg_Error[$eUnknownError]))			;"Translation error" 	"An unknown error occurred."
			SetError(11)
			Return SetError(11, 0, "")
		EndIf

		Local $sGoogleAPIErrorMessage = ""
		$sGoogleAPIErrorMessage = @CRLF & "Google API: "
		If $sGoogleErrorCode <> "" Then
			$sGoogleAPIErrorMessage &=	$sGoogleErrorCode & ": "
		EndIf
		$sGoogleAPIErrorMessage &= $sGoogleError
		MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_Error[$eTransErrorTitle]), $sGoogleAPIErrorMessage)									;"Translation error"	Google API Code and Message
        Return SetError(10, 0, "")
    EndIf

	If $sTranslatedText == "" Then
		MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), _GetStrRes($asArrayMsg_Error[$eTransErrorTitle]), _GetStrRes($asArrayMsg_Error[$eEmptyError]))				;"Translation error"	"The query result is empty."
		Return SetError(12, 0, "")
	EndIf
	;---------------------------------------------------------------
	Local $sDecodeTranslateParsedTextTo_UTF8 = _WinAPI_MultiByteToWideChar($sTranslatedText, 65001, 0, True)
	$sTranslateOutText = StringReplace($sDecodeTranslateParsedTextTo_UTF8, @LF, @CRLF)
	$sTranslateOutText = StringReplace($sTranslateOutText, '\"', '"')
	$sTranslateOutText = StringReplace($sTranslateOutText, '\\', '\')
	Return $sTranslateOutText
EndFunc ;==>_SendTo_Translate_Service
;------------------------------------------------------------------DETECTING DARK-LIGHT THEME
Func _Detect_Dark_Light_Theme()
    Local $sHK = (@OSArch = "X64") ? "HKCU64" : "HKCU"
    Local $sKeyName = $sHK & "\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"

    $nAppsUseLightTheme = RegRead($sKeyName, "AppsUseLightTheme")
    $nSystemUsesLightTheme = RegRead($sKeyName, "SystemUsesLightTheme")
EndFunc
;----------------------------------------------------ACTIVITY OF CONTEXT MENU ITEMS FOR INPUT
Func _ActivityContextMenuItem_Input()
	If _ClipBoard_IsFormatAvailable($CF_UNICODETEXT) Then
		_GUICtrlMenu_SetItemEnabled($hContextMenu_Input, 2) ; Paste
	Else
		_GUICtrlMenu_SetItemDisabled($hContextMenu_Input, 2)
	EndIf
	;-------------------------------------------------
	Local $aSel = _GUICtrlEdit_GetSel($idInput_API_OCRAPIKey)
	If Not IsArray($aSel) Then Return 0

	Local $iSelect = $aSel[1] - $aSel[0]
	If $iSelect = 0 Then
		_GUICtrlMenu_SetItemDisabled($hContextMenu_Input, 0)	;Cut
		_GUICtrlMenu_SetItemDisabled($hContextMenu_Input, 1)	;Copy
		_GUICtrlMenu_SetItemDisabled($hContextMenu_Input, 3)	;Delete
	Else
		_GUICtrlMenu_SetItemEnabled($hContextMenu_Input, 0)		;Cut
		_GUICtrlMenu_SetItemEnabled($hContextMenu_Input, 1)		;Copy
		_GUICtrlMenu_SetItemEnabled($hContextMenu_Input, 3)		;Delete
	EndIf
	;--------------------------------------------------
	Local $iAll = StringLen(GUICtrlRead($idInput_API_OCRAPIKey))
	If $iAll = 0 Or $iAll = $iSelect Then
		_GUICtrlMenu_SetItemDisabled($hContextMenu_Input, 5)	;Select All
	Else
		_GUICtrlMenu_SetItemEnabled($hContextMenu_Input, 5)		;Select All
	EndIf
EndFunc
;-----------------------------------------------------ACTIVITY OF CONTEXT MENU ITEMS FOR EDIT
Func _ActivityContextMenuItem_Edit()
	Local $aSel = _GUICtrlEdit_GetSel($idEdit_RSLT)
	If Not IsArray($aSel) Then Return 0

	Local $iSelect = $aSel[1] - $aSel[0]
	If $iSelect = 0 Then
		_GUICtrlMenu_SetItemDisabled($hContextMenu_Edit, 0)		;Copy
	Else
		_GUICtrlMenu_SetItemEnabled($hContextMenu_Edit, 0)		;Copy
	EndIf
	;-------------------------------------------------
	Local $iAll = StringLen(GUICtrlRead($idEdit_RSLT))
	If $iAll = 0 Or $iAll = $iSelect Then
		_GUICtrlMenu_SetItemDisabled($hContextMenu_Edit, 2)		;Select All
	Else
		_GUICtrlMenu_SetItemEnabled($hContextMenu_Edit, 2)		;Select All
	EndIf
EndFunc
;-------------------------------------------------------------------GET SYSTEM PROXY FOR CURL
Func _GetSystemProxyForCurl()
    Local $sRegPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings"

    Local $bProxyEnable = RegRead($sRegPath, "ProxyEnable")
    If @error Or  Not $bProxyEnable Then Return ""

    Local $sProxyServer = RegRead($sRegPath, "ProxyServer")
    If $sProxyServer = "" Then Return ""

    Return ' --proxy "' & $sProxyServer & '"'
EndFunc
;-------------------------------------------------------------------------CLEARING TASK QUEUE
Func _Clearing_Task_Queue()
    While TrayGetMsg() <> 0
    WEnd

    $bIsProcessing = False
EndFunc
;----------------------------------------------------------WAIT FOR CURL AND ANIMATE THE ICON
Func _Curl_WaitAndAnimate($iPID, $sTitleTimeout, $sMsgTimeout, $sTitleCurlErr, $iErrorOffset = 0)
    Local $hTimer = TimerInit()
    Local $hAnimTimer = TimerInit()
    Local $bAnimState = False ; True
    Local $sStdOut = "", $sStdErr = ""

    ; We define icon sets depending on the system theme
    Local $iBaseIcon = $nSystemUsesLightTheme ? 202 : 201
    Local $iActiveIcon = $nSystemUsesLightTheme ? 204 : 203

	Local $iLastSetIcon = $iBaseIcon

    While 1

		; ------------------Icon Animation
        If $bIconAnimation Then
            Local $iCurrentInterval = $bAnimState ? 200 : 1200

            If TimerDiff($hAnimTimer) > $iCurrentInterval Then
                $bAnimState = Not $bAnimState
                $iLastSetIcon = $bAnimState ? $iActiveIcon : $iBaseIcon
                TraySetIcon(@ScriptFullPath, $iLastSetIcon)
                $hAnimTimer = TimerInit()
            EndIf
        EndIf

        ; ------------Reading Data Streams
        Local $sReadOut = StdoutRead($iPID)
        Local $iErrOut = @error
        Local $sReadErr = StderrRead($iPID)
        Local $iErrErr = @error

        $sStdOut &= $sReadOut
        $sStdErr &= $sReadErr

        ;If both streams are closed, the process is complete.
        If $iErrOut And $iErrErr Then ExitLoop

        ; -------------------------TimeOut
        If TimerDiff($hTimer) > 30000 Then
            ProcessClose($iPID)
            If $iLastSetIcon <> $iBaseIcon Then TraySetIcon(@ScriptFullPath, $iBaseIcon) ; Return the standard icon
            MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), $sTitleTimeout, $sMsgTimeout)

            ; Return an error (2 for OCR or 8 for translation)
            Return SetError(2 + $iErrorOffset, 0, "")
        EndIf

        Sleep(20)
    WEnd

	; Return the standard icon
	If $iLastSetIcon <> $iBaseIcon Then TraySetIcon(@ScriptFullPath, $iBaseIcon)

    ; -------------------------------------------------------------------------------- cURL errors
    Local $asCurlError = StringRegExp($sStdErr, $sRegexCurlError, $STR_REGEXPARRAYMATCH)
    If Not @error Then
        MsgBox(BitOR($MB_ICONERROR, $MB_SYSTEMMODAL, $MB_TOPMOST), $sTitleCurlErr, $asCurlError[0])

        ; Return an error (3 for OCR or 9 for translation)
        Return SetError(3 + $iErrorOffset, 0, "")
    EndIf

    Return $sStdOut
EndFunc
;----------------------------------------------------------------------------------WM COMMAND
Func _WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
	Local $nNotifyCode = BitShift($wParam, 16) ; Get the notification code
	Local $nID = BitAND($wParam, 0xFFFF)       ; Get the control ID
	Local $hCtrl = $lParam                     ; Get the control handle

	If $nID = $idInput_API_OCRAPIKey And $nNotifyCode = $EN_CHANGE Then
		_Checking_Changes_API_Key()
	EndIf

	; Compare the message handle with the input handle by @Mat
	If $hCtrl = $hInput_HK_HotKeyOCR And $nNotifyCode = $EN_CHANGE Then
		_Checking_Changes_Hotkeys()
	EndIf

	If $hCtrl = $hInput_HK_HotKeyTrans And $nNotifyCode = $EN_CHANGE Then
		_Checking_Changes_Hotkeys()
	EndIf

	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_COMMAND
;----------------------------------------------------------CHECK FOR CHANGES IN INPUT API KEY
Func _Checking_Changes_API_Key()
    Local $sCurrentInput = GUICtrlRead($idInput_API_OCRAPIKey)
    Local $iButtonState = GUICtrlGetState($idButton_API_Save)

    If $sCurrentInput <> $sOCRAPIKey Then
        ; If the text has changed AND the button is off, turn it on
        If BitAND($iButtonState, $GUI_DISABLE) Then
            GUICtrlSetState($idButton_API_Save, $GUI_ENABLE)
        EndIf
    Else
        ; If the text is the same as the original AND the button is NOT disabled, disable it.
        If Not BitAND($iButtonState, $GUI_DISABLE) Then
            GUICtrlSetState($idButton_API_Save, $GUI_DISABLE)
        EndIf
    EndIf
EndFunc   ;==>_Checking_Changes_API_Key
;----------------------------------------------------------CHECK FOR CHANGES IN INPUT HOTKEYS
Func _Checking_Changes_Hotkeys()
    Local $sCurrentInput_HK_HotKeyOCR = _GUICtrlHotkey_GetHotkeyCode($hInput_HK_HotKeyOCR)
	Local $sCurrentInput_HK_HotKeyTrans = _GUICtrlHotkey_GetHotkeyCode($hInput_HK_HotKeyTrans)
    Local $iButtonState = GUICtrlGetState($idButton_HK_Save)

    If $sCurrentInput_HK_HotKeyOCR <> $iOCRHotkeyCode Or $sCurrentInput_HK_HotKeyTrans <> $iTransHotkeyCode Then
        ; If the text has changed AND the button is off, turn it on
        If BitAND($iButtonState, $GUI_DISABLE) Then
            GUICtrlSetState($idButton_HK_Save, $GUI_ENABLE)
        EndIf
    Else
        ; If the text is the same as the original AND the button is NOT disabled, disable it.
        If Not BitAND($iButtonState, $GUI_DISABLE) Then
            GUICtrlSetState($idButton_HK_Save, $GUI_DISABLE)
        EndIf
    EndIf
EndFunc   ;==>_Checking_Changes_Hotkeys
;-------------------------------------------------------------------------CHECK GOOGLE ERRORS
Func _Check_Google_Error($mJson)
    Local $sMsg = _JSON_Get($mJson, "error.message")
    If @error Then Return SetError(44, 0, 0)

    $sGoogleError = $sMsg ; write in Global
	;----------------------
    Local $iCode = _JSON_Get($mJson, "error.code")
    If Not @error Then $sGoogleErrorCode = $iCode ; write in Global

    Return 1
EndFunc