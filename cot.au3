#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ico\white-be.ico
#AutoIt3Wrapper_Outfile=COT.exe
#AutoIt3Wrapper_Res_Description=Utility for recognizing and translating text captured from the screen...
#AutoIt3Wrapper_Res_Fileversion=0.9.0.50
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=Capture-OCR-Translate
#AutoIt3Wrapper_Res_ProductVersion=0.9.0
#AutoIt3Wrapper_Res_CompanyName=NyBumBum
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © NyBumBum 2025
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Icon_Add=ico\white-24-16.ico
#AutoIt3Wrapper_Res_Icon_Add=ico\black-24-16.ico
#AutoIt3Wrapper_Res_File_Add=loc\en.ini, 6
#AutoIt3Wrapper_Res_File_Add=loc\ru.ini, 6
#AutoIt3Wrapper_Res_File_Add=cur\cross.cur, 1, 101
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;Hey! Don't forget to replace the cursor in Resource Hacker....
;Don't forget to change the version in About....


#include <Array.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <Constants.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <GDIPlus.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiEdit.au3>
#include <GuiHotkey.au3> ;(UDF by Mat)
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

;-----------------------------------------------------------------------CHECKING IF ALREADY RUNNING
if _Singleton("Capture_OCR_Translate", 1) = 0 Then
    WinActivate("[CLASS:AutoIt v3 GUI;TITLE:Capture-OCR-Translate]")
	Exit
EndIf
;-----------------------------------------------------------------------
Opt("GUICloseOnESC", 0)
Opt('GUIEventOptions', 1)	;for Saving Window Position on minimize
Opt("TrayOnEventMode", 1)
Opt("TrayMenuMode", 1+2+4)

Global Const $sRegexCurlError = '(?m)^(curl:\s\([0-9]{1,2}\).*)$'
Global $bNeedAPIKey = False
Global $idSelectAll	; for castom menu
Global $hEdit, $hInput
Global $hCursorArrow, $hMyCursor

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

#Region;======================== LOCALIZATION (BASED ON FAQ BY YASHIED) ==================================
;----------------------------------------------------------------------EXTRACT STRINGS FROM A STRING TABLE
Func _GetStrRes($iStringID)
	Local $iString = _WinAPI_LoadString($hInstance, $iStringID)
	If @error Then
		MsgBox($MB_ICONERROR, "Error", "Failed to get string from program resources.")
		Exit
	EndIf
	Return $iString
EndFunc   ;==>_GetStringFromResources
;------------------------------------
Global Enum $eAppName, $eTabResult, $eTabGeneral, $eTabAPIKey, $eTabHotkey, $eTabLanguage, $eTabAbout
Global $asArrayGUI_App[7] = [6000, 6001, 6002, 6003, 6004, 6005, 6006]
Global Enum $eClipboard, $eResultOut, $eAlwaysOnTop, $eTrayIconAction
Global $asArrayTabGeneral[4] = [6016, 6017, 6018, 6019]
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
Global Enum $eError, $eTimeoutError, $eUnknown, $eTextNotFound
Global $asArrayMsg_Std[4] = [6144, 6145, 6146, 6147]
Global Enum $eAPI_Success, $eAPI_ProKey, $eAPI_NonCorrectKey, $eAPI_Failed
Global $asArrayMsg_API[4] = [6160, 6161, 6162, 6163]
Global Enum $eImageTooLarge
Global $asArrayMsg_Image[1] = [6176]
Global Enum $eOCR_Removed, $eOCR_Saved, $eOCR_Changed, $eTrans_Removed, $eTrans_Saved, $eTrans_Changed, $eOCRFailed_Reg, $eTransFailed_Reg, $eOCRFailed_Unreg, $eTransFailed_Unreg, $eOCR_Occupied, $eSame
Global $asArrayMsg_Hotkey[12] = [6192, 6193, 6194, 6195, 6196, 6197, 6198, 6199, 6200, 6201, 6202, 6203]
Global Enum $e_curl_Download, $e_curl_OCR, $e_curl_Translate
Global $asArrayMsg_curl[3] = [6208, 6209, 6210]
Global Enum $eDownloadError, $eFailedDownload
Global $asArrayMsg_ListLng[2] = [6224, 6225]
Global Enum $eOCRError, $eOCRParsingError, $eOCREmpty, $eOCRTimeout
Global $asArrayMsg_OCR[4] = [6240, 6241, 6242, 6243]
Global Enum $eTransError, $eTransParsingError, $eTransEmpty, $eTransTimeout
Global $asArrayMsg_Translate[4] = [6256, 6257, 6258, 6259]
#EndRegion;---------------------------------------------------------

Global $sIni_Patch = @ScriptDir & "\settings.ini"
Global $iWidth_GUI = 490, $iHeight_GUI = 300
Global $iWinPos_X, $iWinPos_Y, $sFontResult, $iFontSizeResult, $sPutResultToClipboard, $sDisplayResultToResultTab, $sAlwaysOnTop, $iTrayIconAction, $sOCRAPIKey, $sTranslateAPIKey, $sOCRHotkey, $iOCRHotkeyCode, $sTranslateHotkey, $iTransHotkeyCode, $sTargetLanguage
Global $bPutResultToClipboard, $bDisplayResultToResultTab, $bAlwaysOnTop
Global $aisArrayLanguageCodes[191][2] = [[4,"zh-CN"],[1025,"ar"],[1026,"bg"],[1027,"ca"],[1028,"zh-TW"],[1029,"cs"],[1030,"da"],[1031,"de"],[1032,"el"],[1033,"en"],[1034,"es"],[1035,"fi"],[1036,"fr"],[1037,"he"],[1038,"hu"],[1039,"is"],[1040,"it"],[1041,"ja"],[1042,"ko"],[1043,"nl"],[1044,"no"],[1045,"pl"],[1046,"pt"],[1048,"ro"],[1049,"ru"],[1050,"hr"],[1051,"sk"],[1052,"sq"],[1053,"sv"],[1054,"th"],[1055,"tr"],[1056,"ur"],[1057,"id"],[1058,"uk"],[1059,"be"],[1060,"sl"],[1061,"et"],[1062,"lv"],[1063,"lt"],[1064,"tg"],[1065,"fa"],[1066,"vi"],[1067,"hy"],[1068,"az"],[1069,"eu"],[1071,"mk"],[1074,"tn"],[1076,"xh"],[1077,"zu"],[1078,"af"],[1079,"ka"],[1081,"hi"],[1082,"mt"],[1086,"ms"],[1087,"kk"],[1088,"ky"],[1089,"sw"],[1090,"tk"],[1091,"uz"],[1092,"tt"],[1093,"bn"],[1094,"pa"],[1095,"gu"],[1096,"or"],[1097,"ta"],[1098,"te"],[1099,"kn"],[1100,"ml"],[1101,"as"],[1102,"mr"],[1103,"sa"],[1104,"mn"],[1106,"cy"],[1107,"km"],[1108,"lo"],[1110,"gl"],[1111,"gom"],[1115,"si"],[1118,"am"],[1121,"ne"],[1122,"fy"],[1123,"ps"],[1124,"tl"],[1125,"dv"],[1128,"ha"],[1130,"yo"],[1131,"qu"],[1132,"nso"],[1133,"ba"],[1134,"lb"],[1136,"ig"],[1139,"ti"],[1141,"haw"],[1150,"br"],[1152,"ug"],[1153,"mi"],[1154,"oc"],[1155,"co"],[1159,"rw"],[1169,"gd"],[1170,"ku"],[2049,"ar"],[2051,"ca"],[2052,"zh-CN"],[2055,"de"],[2057,"en"],[2058,"es"],[2060,"fr"],[2064,"it"],[2067,"nl"],[2068,"no"],[2070,"pt-PT"],[2077,"sv"],[2080,"ur"],[2092,"az"],[2098,"tn"],[2108,"ga"],[2110,"ms"],[2117,"bn"],[2118,"pa-Arab"],[2121,"ta"],[2128,"mn"],[2137,"sd"],[2151,"ff"],[2155,"qu"],[2163,"ti"],[3073,"ar"],[3076,"zh-HK"],[3079,"de"],[3081,"en"],[3082,"es"],[3084,"fr"],[3098,"sr"],[3179,"qu"],[4097,"ar"],[4100,"zh-CN"],[4103,"de"],[4105,"en"],[4106,"es"],[4108,"fr"],[4122,"hr"],[5121,"ar"],[5124,"zh-TW"],[5127,"de"],[5129,"en"],[5130,"es"],[5132,"fr"],[5146,"bs"],[6145,"ar"],[6153,"en"],[6154,"es"],[6156,"fr"],[7169,"ar"],[7177,"en"],[7178,"es"],[7194,"sr"],[8193,"ar"],[8201,"en"],[8202,"es"],[9217,"ar"],[9225,"en"],[9226,"es"],[10241,"ar"],[10249,"en"],[10250,"es"],[10266,"sr"],[11265,"ar"],[11273,"en"],[11274,"es"],[12289,"ar"],[12297,"en"],[12298,"es"],[12314,"sr"],[13313,"ar"],[13321,"en"],[13322,"es"],[14337,"ar"],[14346,"es"],[15361,"ar"],[15370,"es"],[16385,"ar"],[16393,"en"],[16394,"es"],[17417,"en"],[17418,"es"],[18441,"en"],[18442,"es"],[19466,"es"],[20490,"es"],[21514,"es"],[31748,"zh-TW"]]

_FixAccelHotKeyLayout()
_Check_Exists_INI_File()
_GUI_Window_Position()
$wProcHandle = DllCallbackRegister("_WindowProc", "ptr", "hwnd;uint;wparam;lparam")	;for custom menu

;backup original arrow cursor
$hCursorArrow = _WinAPI_LoadImage(0, $OCR_NORMAL, $IMAGE_CURSOR, 0, 0, BitOR($LR_DEFAULTSIZE, $LR_SHARED))
$hMyCursor = _WinAPI_LoadImage($hInstance, 101, $IMAGE_CURSOR, 0, 0, $LR_DEFAULTSIZE)

#Region;============================== GUI ==================================================
$hGUI_App = GUICreate(_GetStrRes($asArrayGUI_App[$eAppName]), $iWidth_GUI, $iHeight_GUI, $iWinPos_X, $iWinPos_Y);"Capture-OCR-Translate"
If $nSystemUsesLightTheme Then
	GUISetIcon(@ScriptFullPath, 202)
Else
	GUISetIcon(@ScriptFullPath, 201)
EndIf
GUISetFont(10, 400, 0, "Segoe UI")
$hTab = GUICtrlCreateTab(10, 10, 470, 280)
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
;----------------------------------------------------------------------TAB RESULT
$hTabResult = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabResult]))												;"Result"
$idEdit_RSLT = GUICtrlCreateEdit("", 30, 48, 430, 228, BitOR($ES_AUTOVSCROLL,$ES_READONLY,$ES_WANTRETURN,$WS_VSCROLL), $WS_EX_STATICEDGE)
GUICtrlSetFont(-1, $iFontSizeResult, 400, 0, $sFontResult)

;------------------------------------------Custom Context Menu For Edit
$hContextMenu_Edit = _GUICtrlMenu_CreatePopup()
If $hContextMenu_Edit <> 0 Then
	_GUICtrlMenu_AddMenuItem($hContextMenu_Edit, _GetStrRes($asArrayContextMenu[$eCopy]), $WM_COPY)                			;"Copy"
	_GUICtrlMenu_AddMenuItem($hContextMenu_Edit, "")
	_GUICtrlMenu_AddMenuItem($hContextMenu_Edit, _GetStrRes($asArrayContextMenu[$eSelect_All]), $idSelectAll)       		;"Select All"
	$hEdit = GUICtrlGetHandle($idEdit_RSLT)
	$wProcOld_Edit = _WinAPI_SetWindowLong($hEdit, $GWL_WNDPROC, DllCallbackGetPtr($wProcHandle))
EndIf
;----------------------------------------------------------------------TAB GENERAL
$hTabGeneral = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabGeneral]))												;"General"
$idCheckbox_GNR_Clipboard = GUICtrlCreateCheckbox(_GetStrRes($asArrayTabGeneral[$eClipboard]), 34, 48, 423, 21)				;"Put the result in the clipboard"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
If $sPutResultToClipboard == "True" Then
	$bPutResultToClipboard = True
	GUICtrlSetState(-1, $GUI_CHECKED)
Else
	$bPutResultToClipboard = False
EndIf

$idCheckbox_GNR_ResultOut = GUICtrlCreateCheckbox(_GetStrRes($asArrayTabGeneral[$eResultOut]), 34, 70, 423, 21)				;"Display the result on the Result tab"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
If $sDisplayResultToResultTab == "True" Then
	$bDisplayResultToResultTab = True
	GUICtrlSetState(-1, $GUI_CHECKED)
Else
	$bDisplayResultToResultTab = False
EndIf

$idCheckbox_GNR_AlwaysOnTop = GUICtrlCreateCheckbox(_GetStrRes($asArrayTabGeneral[$eAlwaysOnTop]), 34, 92, 423, 21)			;"Always on Top"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
If $sAlwaysOnTop == "True" Then
	$bAlwaysOnTop = True
	GUICtrlSetState(-1, $GUI_CHECKED)
	WinSetOnTop($hGUI_App, "", 1)
Else
	$bAlwaysOnTop = False
EndIf

$idLabel_GNR_TrayIconAction = GUICtrlCreateLabel(_GetStrRes($asArrayTabGeneral[$eTrayIconAction]), 34, 120, 230, 51, $SS_RIGHT)	;"Clicking on the tray icon launches:"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
$idCombo_GNR_TrayIconAction = GUICtrlCreateCombo("", 277, 118, 180, 25, BitOR($GUI_SS_DEFAULT_COMBO,$CBS_SIMPLE))
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
If $iTrayIconAction < 0 Or $iTrayIconAction > (UBound($asArrayTrayIconAction) - 1) Then
	$iTrayIconAction = 0; Default
	IniWrite($sIni_Patch, "General", "TrayIconAction", 0)
EndIf
GUICtrlSetData(-1, _Fill_Combobox_TrayIconAction(), _GetStrRes($asArrayTrayIconAction[$iTrayIconAction]))

;----------------------------------------------------------------------TAB API KEY
$hTabAPIkey = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabAPIKey]))												;"API Key"
$idLabel_API_Description = GUICtrlCreateLabel(_GetStrRes($asArrayTabAPIKey[$eDescription]), 34, 48, 423, 61)				;"The program requires an API key to work. Get a free API key from the OCR service website and paste the API key in the field below. All you need is a valid email address and 2-3 minutes."
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
$idButton_API_GetOCRAPIKey = GUICtrlCreateButton(_GetStrRes($asArrayTabAPIKey[$eGetOCRAPIKey]), 34, 116, 205, 31)			;"Get API Key"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
$idInput_API_OCRAPIKey = GUICtrlCreateInput("", 252, 119, 205, 25, BitOR($GUI_SS_DEFAULT_INPUT,$ES_RIGHT))
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
If Not ($sOCRAPIKey == "") Then
	GUICtrlSetData(-1, $sOCRAPIKey)
Else
	$bNeedAPIKey = True
EndIf
;---------------------------------------------------Custom Context Menu For Input
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
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
;----------------------------------------------------------------------TAB HOTKEYS
$hTabHotkeys = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabHotkey]))												;"Hotkeys"
$idLabel_HK_HotKeyOCR = GUICtrlCreateLabel(_GetStrRes($asArrayTabHotkey[$eOCRHotkey]), 34, 52, 210, 21, $SS_RIGHT)			;"Hotkey for OCR:"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
$hInput_HK_HotKeyOCR = _GUICtrlHotkey_Create($hGUI_App, 257, 50, 200, 25)
_GUICtrlHotkey_SetRules($hInput_HK_HotKeyOCR, BitOR($HKCOMB_S, $HKCOMB_NONE), $HOTKEYF_ALT)
_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
$sOCRHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyOCR)
If $iOCRHotkeyCode <> 0 Then
	_Registration_Hotkey_OCR()
EndIf
If @error Then
	$iOCRHotkeyCode = 0
	_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
	IniWrite($sIni_Patch, "Hotkey", "OCRHotkeyCode", $iOCRHotkeyCode)
	$bNeedHotKey = True
EndIf
ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
;-------------
$idLabel_HK_HotKeyTrans = GUICtrlCreateLabel(_GetStrRes($asArrayTabHotkey[$eTransHotkey]), 34, 88, 210, 21, $SS_RIGHT)		;"Hotkey for translation:"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
$hInput_HK_HotKeyTrans = _GUICtrlHotkey_Create($hGUI_App, 257, 86, 200, 25)
_GUICtrlHotkey_SetRules($hInput_HK_HotKeyTrans, BitOR($HKCOMB_S, $HKCOMB_NONE), $HOTKEYF_ALT)

_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
$sTranslateHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyTrans)
If $iTransHotkeyCode <> 0 Then
	_Registration_Hotkey_Translate()
EndIf
If @error Then
	$iTransHotkeyCode = 0
	_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
	IniWrite($sIni_Patch, "Hotkey", "TranslateHotkeyCode", $iTransHotkeyCode)
	$bNeedHotKey = True
EndIf
ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
$idLabel_HK_Tip = GUICtrlCreateLabel(_GetStrRes($asArrayTabHotkey[$eTipHotkey]), 34, 121, 423, 81)							;"Single characters and Shift-only characters are blocked to prevent software functions from running while typing. Please show a little imagination..."
;--------------
$idLabel_HK_Hr = GUICtrlCreateLabel("", 11, 238, 466, 1)
GUICtrlSetBkColor(-1, 0xD9D9D9)
$idButton_HK_Save = GUICtrlCreateButton(_GetStrRes($asArrayTabAPIKey[$eSave]), 357, 248, 100, 31)							;"Save"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
;----------------------------------------------------------------------TAB LANGUAGE
$hTabLanguage = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabLanguage]))											;"Language"
$idLabel_LNG_TargetLanguage = GUICtrlCreateLabel(_GetStrRes($asArrayTabLanguage[$eTargetLanguage]), 34, 48, 210, 41, $SS_RIGHT)
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
$idCombo_LNG_TargetLanguage = GUICtrlCreateCombo("", 257, 54, 200, 25, BitOR($GUI_SS_DEFAULT_COMBO,$CBS_SIMPLE,$CBS_SORT))	;"Target language for translating recognized text:"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
Local $sGoogleLangID = _GET_GoogletLngID_from_SystemLngID()
Local $asArrayLanguageCodeAndName = _Download_Translation_Language_List($sGoogleLangID)
If @error Then
	Exit
EndIf
If $sTargetLanguage == "" Then
	$sTargetLanguage = $sGoogleLangID
	IniWrite($sIni_Patch, "Language", "TargetLanguage", $sTargetLanguage)
EndIf
GUICtrlSetData(-1, _Fill_Combobox_TargetLanguage(), _Get_Name_TargetLanguage())
;----------------------------------------------------------------------TAB ABOUT
$hTabAbout = GUICtrlCreateTabItem(_GetStrRes($asArrayGUI_App[$eTabAbout]))													;"About"
$idIcon_ABT = GUICtrlCreateIcon(@ScriptFullPath, 99, 404, 44, 48, 48)
$idLabel_ABT_AppName = GUICtrlCreateLabel(_GetStrRes($asArrayGUI_App[$eAppName]), 34, 52, 323, 41)							;"Capture-OCR-Translate"
GUICtrlSetFont(-1, 20, 800, 0, "Segoe UI")
GUICtrlSetColor(-1, 0xD9D9D9)
$idLabel_ABT_Version = GUICtrlCreateLabel("0.9.0", 34, 110, 350, 21)														;"X.X.X"----------------------------------->>>>>>>       !
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
$idLabel_ABT_Copyright = GUICtrlCreateLabel(_GetStrRes($asArrayTabAbout[$eCopyright]), 34, 131, 350, 21)					;"Copyright © NyBumBum 2025"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
$idLabel_ABT_Mail = GUICtrlCreateLabel("nybumbum@gmail.com", 34, 152, 350, 21)												;"nybumbum@gmail.com"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x0078D7)
GUICtrlSetCursor (-1, 0)
$idLabel_ABT_Site = GUICtrlCreateLabel("github.com/nbb1967/capture-ocr-translate", 34, 173, 350, 21)						;"github.com/nbb1967/capture-ocr-translate"
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x0078D7)
GUICtrlSetCursor (-1, 0)
$idLabel_ABT_Components = GUICtrlCreateLabel(_GetStrRes($asArrayTabAbout[$eComponents]), 34, 210, 420, 21)					;"Based on AutoIt, curl, OCR.SPACE and Google Translation."
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
$idLabel_ABT_Warranty = GUICtrlCreateLabel(_GetStrRes($asArrayTabAbout[$eWarranty]), 34, 230, 422, 41)						;'This free utility is provided "as is" without any warranty...'
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
$Label1 = GUICtrlCreateLabel("", 11, 98, 466, 1)
GUICtrlSetFont(-1, 10, 400, 0, "Segoe UI")
GUICtrlSetBkColor(-1, 0xD9D9D9)
GUICtrlCreateTabItem("")

If $bNeedAPIKey Then
	GUISetState(@SW_SHOW, $hGUI_App)
	GUICtrlSetState($hTabAPIkey, $GUI_SHOW)
	$iLastTab = 2
Else
	GUISetState(@SW_HIDE, $hGUI_App)
	GUICtrlSetState($hTabResult, $GUI_SHOW)
	$iLastTab = 0
EndIf
#EndRegion

#Region;==============================TRAY GUI==============================================
$idTrayMenu_Translate = TrayCreateItem(_GetStrRes($asArrayGUI_Tray[$eTray_Translate]))									;"Translate"
TrayItemSetOnEvent(-1, "_OCR_and_Translate_Tray_Function")
$idTrayMenu_OCR = TrayCreateItem(_GetStrRes($asArrayGUI_Tray[$eTray_OCR]))												;"OCR"
TrayItemSetOnEvent(-1, "_Only_OCR_Tray_Function")
TrayCreateItem("")
$idTrayMenu_Result = TrayCreateItem(_GetStrRes($asArrayGUI_Tray[$eTray_Result]))										;"Result"
TrayItemSetOnEvent(-1, "_Result_Tray_Function")
TrayCreateItem("")
$idTrayMenu_Exit = TrayCreateItem(_GetStrRes($asArrayGUI_Tray[$eExit]))													;"Exit"
TrayItemSetOnEvent(-1, "_Exit_Tray_Function")
TraySetOnEvent($TRAY_EVENT_PRIMARYDOWN, "_Tray_Icon_Action")
If $nSystemUsesLightTheme Then
	TraySetIcon(@ScriptFullPath, 202)
Else
	TraySetIcon(@ScriptFullPath, 201)
EndIf

TraySetClick($TRAY_CLICK_SECONDARYDOWN)
TraySetToolTip (_GetStrRes($asArrayGUI_App[$eAppName]))																	;"Capture-OCR-Translate"

_Tray_Icon_Action_Item_Select()

;GUISetState(@SW_SHOW)
#EndRegion


Local $aGUIWindowPosition
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			;--------------------------------------------------------- Save Window Position
			$aGUIWindowPosition = WinGetPos($hGUI_App)
			If $aGUIWindowPosition[0] <> -32000 And $aGUIWindowPosition[1] <> -32000 Then
				IniWrite($sIni_Patch, "GUI", "GUIWindowPosition_X", $aGUIWindowPosition[0])
				IniWrite($sIni_Patch, "GUI", "GUIWindowPosition_Y", $aGUIWindowPosition[1])
			EndIf
			GUISetState(@SW_HIDE, $hGUI_App)
		Case $GUI_EVENT_MINIMIZE
			$aGUIWindowPosition = WinGetPos($hGUI_App)
			GUISetState(@SW_MINIMIZE, $hGUI_App)
			IniWrite($sIni_Patch, "GUI", "GUIWindowPosition_X", $aGUIWindowPosition[0])
			IniWrite($sIni_Patch, "GUI", "GUIWindowPosition_Y", $aGUIWindowPosition[1])
        Case $GUI_EVENT_RESTORE
            GUISetState(@SW_RESTORE, $hGUI_App)
        Case $GUI_EVENT_MAXIMIZE
            GUISetState(@SW_MAXIMIZE, $hGUI_App)
		Case $hTab	;----------------------------------------------------------------- Tab
			$iCurrentTab = GUICtrlRead($hTab)
			If $iCurrentTab <> $iLastTab Then
				Switch $iCurrentTab
					Case 0
						ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
						ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
					Case 1
						ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
						ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
					Case 2
						ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
						ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
					Case 3
						ControlShow($hGUI_App, "", $hInput_HK_HotKeyOCR)
						ControlShow($hGUI_App, "", $hInput_HK_HotKeyTrans)
					Case 4
						ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
						ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
					Case 5
						ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
						ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
				EndSwitch
				$iLastTab = $iCurrentTab
			EndIf

		Case $idCheckbox_GNR_Clipboard	;---------------------------------------------- Clipboard
			If BitAND(GUICtrlRead($idCheckbox_GNR_Clipboard), $GUI_CHECKED) = $GUI_CHECKED Then
				$bPutResultToClipboard = True
			Else
				$bPutResultToClipboard = False
			EndIf
			IniWrite($sIni_Patch, "General", "PutResultToClipboard", $bPutResultToClipboard)
		Case $idCheckbox_GNR_ResultOut	;----------------------------------------------- Result Out
			If BitAND(GUICtrlRead($idCheckbox_GNR_ResultOut), $GUI_CHECKED) = $GUI_CHECKED Then
				$bDisplayResultToResultTab = True
			Else
				$bDisplayResultToResultTab = False
			EndIf
			IniWrite($sIni_Patch, "General", "DisplayResultToResultTab", $bDisplayResultToResultTab)
		Case $idCheckbox_GNR_AlwaysOnTop	;------------------------------------------- Always on Top
			If BitAND(GUICtrlRead($idCheckbox_GNR_AlwaysOnTop), $GUI_CHECKED) = $GUI_CHECKED Then
				WinSetOnTop($hGUI_App, "", 1)
				$bAlwaysOnTop = True
			Else
				WinSetOnTop($hGUI_App, "", 0)
				$bAlwaysOnTop = False
			EndIf
			IniWrite($sIni_Patch, "General", "AlwaysOnTop", $bAlwaysOnTop)
		Case $idCombo_GNR_TrayIconAction
			If Not (GUICtrlRead($idCombo_GNR_TrayIconAction) == _GetStrRes($asArrayTrayIconAction[$iTrayIconAction])) Then
				For $i = 0 To UBound($asArrayTrayIconAction) - 1
					If GUICtrlRead($idCombo_GNR_TrayIconAction) == _GetStrRes($asArrayTrayIconAction[$i]) Then
						$iTrayIconAction = $i
						IniWrite($sIni_Patch,  "General", "TrayIconAction", $iTrayIconAction)
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
						IniWrite($sIni_Patch, "Language", "TargetLanguage", $sTargetLanguage)
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
WEnd
;=================================================================================================

;----------------------------------------------------------------------CHECKING THAT THE INI FILE EXISTS
Func _Check_Exists_INI_File()
	Local $bSuccess = FileExists ($sIni_Patch)
	If $bSuccess = 1 Then
		_Read_Settings()
	Else
		_Create_INI_File()
		_Read_Settings()
	EndIf

EndFunc
;-----------------------------------------------------------------------CREATE INI FILE
Func _Create_INI_File()
	IniWriteSection($sIni_Patch, "GUI", "GUIWindowPosition_X=" & @LF & "GUIWindowPosition_Y=")
	IniWriteSection($sIni_Patch, "Result", "FontResult=Segoe UI" & @LF & "FontSizeResult=10")
	IniWriteSection($sIni_Patch, "General", "PutResultToClipboard=True" & @LF & "DisplayResultToResultTab=True" & @LF & "AlwaysOnTop=False" & @LF & "TrayIconAction=0")
	IniWriteSection($sIni_Patch, "APIKey", "OCRAPIKey=" & @LF & "TranslateAPIKey=AIzaSyBOti4mM-6x9WDnZIjIeyEU21OpBXqWBgw")
	IniWriteSection($sIni_Patch, "Hotkey", "OCRHotkeyCode=" & @LF & "TranslateHotkeyCode=")
	IniWriteSection($sIni_Patch, "Language", "TargetLanguage=")
EndFunc
;----------------------------------------------------------------------READING USER SETTINGS FROM INI FILE
Func _Read_Settings()
	$iWinPos_X = IniRead($sIni_Patch, "GUI", "GUIWindowPosition_X", "")
	$iWinPos_Y = IniRead($sIni_Patch, "GUI", "GUIWindowPosition_Y", "")
	$sFontResult = IniRead($sIni_Patch, "Result", "FontResult", "Segoe UI")
	$iFontSizeResult = IniRead($sIni_Patch, "Result", "FontSizeResult", 10)
	$sPutResultToClipboard = IniRead($sIni_Patch, "General", "PutResultToClipboard", "True")
	$sDisplayResultToResultTab = IniRead($sIni_Patch, "General", "DisplayResultToResultTab", "True")
	$sAlwaysOnTop = IniRead($sIni_Patch, "General", "AlwaysOnTop", "False")
	$iTrayIconAction = IniRead($sIni_Patch, "General", "TrayIconAction", 0)
	$sOCRAPIKey = IniRead($sIni_Patch, "APIKey", "OCRAPIKey", "")
	$sTranslateAPIKey = IniRead($sIni_Patch, "APIKey", "TranslateAPIKey", "AIzaSyBOti4mM-6x9WDnZIjIeyEU21OpBXqWBgw")
	$iOCRHotkeyCode = IniRead($sIni_Patch, "Hotkey", "OCRHotkeyCode", "")
	$iTransHotkeyCode = IniRead($sIni_Patch, "Hotkey", "TranslateHotkeyCode", "")
	$sTargetLanguage = IniRead($sIni_Patch, "Language", "TargetLanguage", "")
EndFunc
;----------------------------------------------------------------------CHECKING THE POSITION OF THE PROGRAM WINDOW. RETURN AN INACCESSIBLE WINDOW
Func _GUI_Window_Position()
	Local $tRECT_WorkArea, $iWidth_WorkArea, $iHeight_WorkArea
	$tRECT_WorkArea = _WinAPI_GetWorkArea()
	$iWidth_WorkArea = DllStructGetData($tRECT_WorkArea, 'Right') - DllStructGetData($tRECT_WorkArea, 'Left')
	$iHeight_WorkArea = DllStructGetData($tRECT_WorkArea, 'Bottom') - DllStructGetData($tRECT_WorkArea, 'Top')

	If Not ($iWinPos_X == "") And Not ($iWinPos_Y == "") Then
		If $iWinPos_X > -$iWidth_GUI + 150 And $iWinPos_X < $iWidth_WorkArea - 50 And $iWinPos_Y > -10 And $iWinPos_Y < $iHeight_WorkArea - 30 Then
			Return
		EndIf
	EndIf

	Local $iHeight_Caption = _WinAPI_GetSystemMetrics($SM_CYCAPTION)
	Local $iSize_Edge = _WinAPI_GetSystemMetrics($SM_CXEDGE)
	$iWinPos_X = $iWidth_WorkArea - $iWidth_GUI - 2 * $iSize_Edge
	$iWinPos_Y = $iHeight_WorkArea - $iHeight_Caption - $iHeight_GUI - 2 * $iSize_Edge
EndFunc ;==> _GUI_Window_Position
;----------------------------------------------------------------------CUSTOM CONTEXT MENU BY RASYM
Func _WindowProc($hWnd, $Msg, $wParam, $lParam)
	Local $aRet
	Switch $hWnd
		Case $hInput
			Switch $Msg
				Case $WM_CONTEXTMENU
					_ActivityContextMenuItem_Input()
					_GUICtrlMenu_TrackPopupMenu($hContextMenu_Input, $wParam)
					Return 1
				Case $WM_COMMAND
					Switch $wParam
						Case $WM_CUT, $WM_COPY, $WM_PASTE, $WM_CLEAR
							_SendMessage($hWnd, $wParam)
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
					_ActivityContextMenuItem_Edit()
					_GUICtrlMenu_TrackPopupMenu($hContextMenu_Edit, $wParam)
					Return 1
				Case $WM_COMMAND
					Switch $wParam
						Case $WM_COPY
							_SendMessage($hWnd, $wParam)
						Case $idSelectAll
							_SendMessage($hWnd, $EM_SETSEL, 0, -1)
					EndSwitch
			EndSwitch
			$aRet = DllCall("user32.dll", "int", "CallWindowProc", "ptr", $wProcOld_Edit, _
					"hwnd", $hWnd, "uint", $Msg, "wparam", $wParam, "lparam", $lParam)
			Return $aRet[0]
	EndSwitch
EndFunc   ;==>_WindowProc
;----------------------------------------------------------------------FILL COMBOBOX TRAY-ICON-ACTION
Func _Fill_Combobox_TrayIconAction()
	Local $sList = ""
	Local $imax = UBound($asArrayTrayIconAction) - 1
	For $i = 0 To $imax
		$sList &= _GetStrRes($asArrayTrayIconAction[$i]) & '|'
	Next
	$sList = StringTrimRight($sList, 1)
	Return $sList

EndFunc   ;==>_Fill_Combobox_TrayIconAction
;----------------------------------------------------------------------CHECK API KEY
Func _Check_OCRAPIKey()
	Local $s_tmp_OCRAPIKey
	Local Const $sRegexOCRAPIKey = '(?i)([K][0-9]{14})'
	Local Const $sRegexOCRAPIKey_Pro = '(?i)([A-Z][A-Z0-9]{12})'
	Local $iOCRAPIKeyLength, $iOCRAPIKeyCorrectness

	$s_tmp_OCRAPIKey = GUICtrlRead($idInput_API_OCRAPIKey)
	If $sOCRAPIKey == $s_tmp_OCRAPIKey Then
		Return
	EndIf

	$iOCRAPIKeyLength = StringLen($s_tmp_OCRAPIKey)
	Switch $iOCRAPIKeyLength
		Case 15
			$iOCRAPIKeyCorrectness = StringRegExp($s_tmp_OCRAPIKey, $sRegexOCRAPIKey)
			If $iOCRAPIKeyCorrectness = 1 Then
				$sOCRAPIKey = $s_tmp_OCRAPIKey
				Local $bSuccess = IniWrite($sIni_Patch, "APIKey", "OCRAPIKey", $sOCRAPIKey)
				If $bSuccess = 1 Then
					MsgBox(0, "", _GetStrRes($asArrayMsg_API[$eAPI_Success]))												;"Free OCR API Key saved successfully."
					$bNeedAPIKey = False
				Else
					MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eError]), _GetStrRes($asArrayMsg_API[$eAPI_Failed]))	;"Error"	"Failed to save Free OCR API Key."
					SetError(14)
				EndIf
				Return
			EndIf
		Case 13
			$iOCRAPIKeyCorrectness = StringRegExp($s_tmp_OCRAPIKey, $sRegexOCRAPIKey_Pro)
			If $iOCRAPIKeyCorrectness = 1 Then
				MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eError]), _GetStrRes($asArrayMsg_API[$eAPI_ProKey]))		;"Error"	"You probably entered a PRO/PRO PDF key. Unfortunately, this program currently only supports Free OCR keys."
				SetError(15)
				Return
			EndIf
	EndSwitch
	MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eError]), _GetStrRes($asArrayMsg_API[$eAPI_NonCorrectKey]))			;"Error"	"Please make sure you have entered a valid API key for Free OCR."
	SetError(16)
	Return
EndFunc
;----------------------------------------------------------------------DOWNLOAD TRANSLATION LANGUAGE LIST
Func _Download_Translation_Language_List($sGoogleLangID)
Local $sCurlCmd = 'curl --ssl-no-revoke -X GET "https://translation.googleapis.com/language/translate/v2/languages?key=' & $sTranslateAPIKey & '&target=' & $sGoogleLangID  & '"'
Local $iPID = Run($sCurlCmd, "", "", BitOR($STDERR_CHILD, $STDOUT_CHILD))
;-----------------------------------------------------------------------
Local $hTimer = TimerInit()
Local $sStdErr = ""
While 1 And Sleep(20)
	$sStdErr &= StderrRead($iPID)
	If @error Then
		ExitLoop
	EndIf
	$fDiff = TimerDiff($hTimer)
	If $fDiff > 30000 Then
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eTimeoutError]), _GetStrRes($asArrayMsg_Translate[$eTransTimeout]))	;"Timeout Error"	"The Translation provider did not respond within 30 seconds."
		SetError(14)
		Return
	EndIf
WEnd
;--------------------------------------------------------------------curl Error
Local $iCurlErrorDetect, $asCurlError
$iCurlErrorDetect = StringRegExp($sStdErr, $sRegexCurlError)
If $iCurlErrorDetect = 1 Then
	$asCurlError = StringRegExp($sStdErr, $sRegexCurlError, $STR_REGEXPARRAYMATCH)
	MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_curl[$e_curl_Download]), $asCurlError[0])		;"curl Error during download"
	SetError(15)
	Return
EndIf
;--------------------------------------------------------------------------------
Local $sStdOut = ""

While 1 And Sleep(20)
	$sStdOut &= StdoutRead($iPID)
	If @error Then
		ExitLoop
	EndIf

WEnd
;------------------------------------------------------------------Parsing List
Local $sJSON_LangugeName = _WinAPI_MultiByteToWideChar($sStdOut, 65001, 0, True)

Local Const $sRegex_JSON_Crop_Start = '(^[^\[]*)'
Local Const $sRegex_JSON_Crop_End = '([^\]]*$)'
$sJSON_LangugeName = StringRegExpReplace($sJSON_LangugeName, $sRegex_JSON_Crop_Start, "")
$sJSON_LangugeName = StringRegExpReplace($sJSON_LangugeName, $sRegex_JSON_Crop_End, "")
Local Const $sRegex_JSON_Split = '({[^\}]*})'
Local $iListDetect = StringRegExp($sJSON_LangugeName, $sRegex_JSON_Split)
If $iListDetect = 1 Then
	Local $asArraySplitList = StringRegExp($sJSON_LangugeName, $sRegex_JSON_Split, $STR_REGEXPARRAYGLOBALMATCH)
	Local $imax = UBound($asArraySplitList, $UBOUND_ROWS) - 1
	Local Const $sRegexLanguageCode = '(?:\"language\":\s\")([^\"]*)'
	Local Const $sRegexLanguageName = '(?:\"name\":\s\")([^\"]*)'
	Local $asArrayLanguageCodeAndName[$imax + 1][2]
	Local $iCodeDetect, $iNameDetect

	For $i = 0 To $imax
		$iCodeDetect = StringRegExp($asArraySplitList[$i], $sRegexLanguageCode)
		If $iCodeDetect = 1 Then
			$as_tmp = StringRegExp($asArraySplitList[$i], $sRegexLanguageCode, $STR_REGEXPARRAYMATCH)
			$asArrayLanguageCodeAndName[$i][0] = $as_tmp[0]
		EndIf
		$iNameDetect = StringRegExp($asArraySplitList[$i], $sRegexLanguageName)
		If $iNameDetect = 1 Then
			$as_tmp = StringRegExp($asArraySplitList[$i], $sRegexLanguageName, $STR_REGEXPARRAYMATCH)
			$asArrayLanguageCodeAndName[$i][1] = $as_tmp[0]
		EndIf

	Next
Else
	MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_ListLng[$eDownloadError]), _GetStrRes($asArrayMsg_ListLng[$eFailedDownload]))		;"Download error"	"Failed to load list of translation languages"
	SetError(16)
	Return

EndIf
Return $asArrayLanguageCodeAndName
EndFunc ;==>_Download_Translation_Language_List
;----------------------------------------------------------------------GET GOOGLE LANGUAGE ID FROM SYSTEM LANGUAGE ID
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
;----------------------------------------------------------------------FILL COMBOBOX TARGET LANGUAGE
Func _Fill_Combobox_TargetLanguage()
	Local $sList = ""
	Local $imax = UBound($asArrayLanguageCodeAndName, $UBOUND_ROWS) - 1
	For $i = 0 To $imax
		$sList &= $asArrayLanguageCodeAndName[$i][1] & '|'
	Next
	$sList = StringTrimRight($sList, 1)
	Return $sList

EndFunc   ;==>_Fill_Combobox_TargetLanguage
;----------------------------------------------------------------------GET NAME TARGET LANGUAGE
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
;----------------------------------------------------------------------START OCR WITH HOTKEY (BASED ON CREATOR)
Func _Only_OCR_Hotkey_Function()
	Local $sHotkey = @HotKeyPressed
	HotKeySet($sHotkey)
	While _IsPressed("10", $hDLL) Or _IsPressed("11", $hDLL) Or _IsPressed("12", $hDLL)
		Sleep(10)
	WEnd

	_Only_OCR_Tray_Function()
	HotKeySet($sHotkey, "_Only_OCR_Hotkey_Function")
EndFunc	;==>_Only_OCR_Hotkey_Function
;----------------------------------------------------------------------START OCR AND TRANSLATION WITH HOTKEY (BASED ON CREATOR)
Func _OCR_and_Translate_Hotkey_Function()
	Local $sHotkey = @HotKeyPressed
	HotKeySet($sHotkey)
	While _IsPressed("10", $hDLL) Or _IsPressed("11", $hDLL) Or _IsPressed("12", $hDLL)
		Sleep(10)
	WEnd

	_OCR_and_Translate_Tray_Function()
	HotKeySet($sHotkey, "_OCR_and_Translate_Hotkey_Function")
EndFunc	;==>_OCR_and_Translate_Hotkey_Function
;----------------------------------------------------------------------FIX ACCEL HOTKEY LAYOUT BY CREATOR
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
;----------------------------------------------------------------------CHECK ALL HOTKEYS
Func _Check_Hotkeys()
	_FixAccelHotKeyLayout()
	Local $i_tmp_OCRHotkeyCode = _GUICtrlHotkey_GetHotkeyCode($hInput_HK_HotKeyOCR)
	Local $i_tmp_TransHotkeyCode = _GUICtrlHotkey_GetHotkeyCode($hInput_HK_HotKeyTrans)

	If $i_tmp_OCRHotkeyCode = $i_tmp_TransHotkeyCode Then
		If $i_tmp_OCRHotkeyCode = 0 Then
			_Check_OCR_HK($i_tmp_OCRHotkeyCode)
			_Check_Trans_HK($i_tmp_TransHotkeyCode)
		Else
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eError]), _GetStrRes($asArrayMsg_Hotkey[$eSame]))					;"Error"	"The hotkeys for two functions cannot be the same."
			_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
			_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
			SetError(20)
			Return
		EndIf
	Else
		If $i_tmp_OCRHotkeyCode = $iTransHotkeyCode And Not ($i_tmp_OCRHotkeyCode = 0) Then
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eError]), _GetStrRes($asArrayMsg_Hotkey[$eOCR_Occupied]))			;"Error"	"The hotkey for OCR is occupied. Try replacing hotkeys one by one."
			_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
			_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
			SetError(21)
		Else
			_Check_OCR_HK($i_tmp_OCRHotkeyCode)
			_Check_Trans_HK($i_tmp_TransHotkeyCode)
		EndIf
	EndIf
EndFunc   ;==>_Check_Hotkeys
;----------------------------------------------------------------------CHECK OCR HOTKEY
Func _Check_OCR_HK($i_tmp_OCRHotkeyCode)
	Select
		Case $i_tmp_OCRHotkeyCode = $iOCRHotkeyCode									;nothing happened
			Return
		Case $i_tmp_OCRHotkeyCode = 0 And $iOCRHotkeyCode <> 0						;remove hotkey
			_Unregistration_Hotkey_OCR()
			If Not @error Then
				$iOCRHotkeyCode = $i_tmp_OCRHotkeyCode
				IniWrite($sIni_Patch, "Hotkey", "OCRHotkeyCode", $iOCRHotkeyCode)
				MsgBox(0, "", _GetStrRes($asArrayMsg_Hotkey[$eOCR_Removed]))								;"Hotkey for OCR successfully removed."
				$sOCRHotkey = ""
			EndIf
			Return
		Case $i_tmp_OCRHotkeyCode <> 0 And $iOCRHotkeyCode = 0						;create hotkey
			_Registration_Hotkey_OCR()
			If Not @error Then
				$iOCRHotkeyCode = $i_tmp_OCRHotkeyCode
				IniWrite($sIni_Patch, "Hotkey", "OCRHotkeyCode", $iOCRHotkeyCode)
				MsgBox(0, "", _GetStrRes($asArrayMsg_Hotkey[$eOCR_Saved]))									;"Hotkey for OCR successfully saved."
				$sOCRHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyOCR)
			EndIf
			Return
		Case $i_tmp_OCRHotkeyCode <> 0 And $iOCRHotkeyCode <> 0						;change hotkey
			_Unregistration_Hotkey_OCR()
			If @error Then
				Return
			Else
				$iOCRHotkeyCode = 0
				$sOCRHotkey = ""
				IniWrite($sIni_Patch, "Hotkey", "OCRHotkeyCode", $iOCRHotkeyCode)
			EndIf
			_Registration_Hotkey_OCR()
			If Not @error Then
				$iOCRHotkeyCode = $i_tmp_OCRHotkeyCode
				IniWrite($sIni_Patch, "Hotkey", "OCRHotkeyCode", $iOCRHotkeyCode)
				MsgBox(0, "", _GetStrRes($asArrayMsg_Hotkey[$eOCR_Changed]))								;"Hotkey for OCR successfully changed."
				$sOCRHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyOCR)
			EndIf
	EndSelect
EndFunc   ;==>_Check_OCR_HK
;----------------------------------------------------------------------CHECK TRANSLATE HOTKEY
Func _Check_Trans_HK($i_tmp_TransHotkeyCode)
	Select
		Case $i_tmp_TransHotkeyCode = $iTransHotkeyCode								;nothing happened
			Return
		Case $i_tmp_TransHotkeyCode = 0 And $iTransHotkeyCode <> 0					;remove hotkey
			_Unregistration_Hotkey_Translate()
			If Not @error Then
				$iTransHotkeyCode = $i_tmp_TransHotkeyCode
				IniWrite($sIni_Patch, "Hotkey", "TranslateHotkeyCode", $iTransHotkeyCode)
				MsgBox(0, "", _GetStrRes($asArrayMsg_Hotkey[$eTrans_Removed]))								;"Hotkey for translation successfully removed."
				$sTranslateHotkey = ""
			EndIf
			Return
		Case $i_tmp_TransHotkeyCode <> 0 And $iTransHotkeyCode = 0					;create hotkey
			_Registration_Hotkey_Translate()
			If Not @error Then
				$iTransHotkeyCode = $i_tmp_TransHotkeyCode
				IniWrite($sIni_Patch, "Hotkey", "TranslateHotkeyCode", $iTransHotkeyCode)
				MsgBox(0, "", _GetStrRes($asArrayMsg_Hotkey[$eTrans_Saved]))								;"Hotkey for translation successfully saved."
				$sTranslateHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyTrans)
			EndIf
			Return
		Case $i_tmp_TransHotkeyCode <> 0 And $iTransHotkeyCode <> 0					;change hotkey
			_Unregistration_Hotkey_Translate()
			If @error Then
				Return
			Else
				$iTransHotkeyCode = 0
				$sTranslateHotkey = ""
				IniWrite($sIni_Patch, "Hotkey", "TranslateHotkeyCode", $iTransHotkeyCode)
			EndIf
			_Registration_Hotkey_Translate()
			If Not @error Then
				$iTransHotkeyCode = $i_tmp_TransHotkeyCode
				IniWrite($sIni_Patch, "Hotkey", "TranslateHotkeyCode", $iTransHotkeyCode)
				MsgBox(0, "", _GetStrRes($asArrayMsg_Hotkey[$eTrans_Changed]))								;"Hotkey for translation successfully changed."
				$sTranslateHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyTrans)
			EndIf
	EndSelect
EndFunc   ;==>_Check_Trans_HK
;----------------------------------------------------------------------REGISTRATION OCR HOTKEY
Func _Registration_Hotkey_OCR()
	Local $s_TMP_OCRHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyOCR)
	Local $bSuccess = HotKeySet(StringLower($s_TMP_OCRHotkey), "_Only_OCR_Hotkey_Function")
	If Not $bSuccess Then
		_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eError]), _GetStrRes($asArrayMsg_Hotkey[$eOCRFailed_Reg]))					;"Error"	"Hotkey for OCR could not be registered in Windows (they may be in use by another program). Try using a different hotkey for OCR."
		SetError(17)
	EndIf
EndFunc   ;==>_Registration_Hotkey_OCR
;----------------------------------------------------------------------REGISTRATION TRANSLATE HOTKEY
Func _Registration_Hotkey_Translate()
	Local $s_TMP_TransHotkey = _GUICtrlHotkey_GetHotkey($hInput_HK_HotKeyTrans)
	Local $bSuccess = HotKeySet(StringLower($s_TMP_TransHotkey), "_OCR_and_Translate_Hotkey_Function")
	If Not $bSuccess Then
		_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eError]), _GetStrRes($asArrayMsg_Hotkey[$eTransFailed_Reg]))					;"Error" 	"Hotkey for translation could not be registered in Windows (they may be in use by another program). Try using a different hotkey for translation."
		SetError(18)
	EndIf
EndFunc   ;==>_Registration_Hotkey_Translate
;----------------------------------------------------------------------UNREGISTRATION OCR HOTKEY
Func _Unregistration_Hotkey_OCR()
	Local $bSuccess = HotKeySet(StringLower($sOCRHotkey))
	If Not $bSuccess Then
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eError]), _GetStrRes($asArrayMsg_Hotkey[$eOCRFailed_Unreg]))					;"Error"	"Failed to unregister hotkey for OCR."
		_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyOCR, $iOCRHotkeyCode)
		SetError(21)
	EndIf
EndFunc   ;==>_Unregistration_Hotkey_OCR
;----------------------------------------------------------------------UNREGISTRATION TRANSLATE HOTKEY
Func _Unregistration_Hotkey_Translate()
	Local $bSuccess = HotKeySet(StringLower($sTranslateHotkey))
	If Not $bSuccess Then
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eError]), _GetStrRes($asArrayMsg_Hotkey[$eTransFailed_Unreg]))				;"Error"	"Failed to unregister hotkey for translation."
		_GUICtrlHotkey_SetHotkeyCode($hInput_HK_HotKeyTrans, $iTransHotkeyCode)
		SetError(22)
	EndIf
EndFunc   ;==>_Unregistration_Hotkey_Translate
;----------------------------------------------------------------------ACTION BY CLICKING ON THE TRAY ICON
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
;----------------------------------------------------------------------SELECTING ITEM MENU FOR ACTION BY CLICKING ON THE TRAY ICON
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
	Local $aiPos = _ScreenRegion_GetRect(0xFF0078D7, 0x460066CC, 1)
	If @error Then
		Return
	EndIf
	Local $hCapturedBitmap = _ScreenCapture_Capture("", $aiPos[0], $aiPos[1], $aiPos[2], $aiPos[3], False)
	Local $sWindowsFullPathToImageFile = _SaveCapturedBitmap($hCapturedBitmap)
	If @error Then
		Return
	EndIf

	Local $sOCR_OutText = _SendTo_OCR_Service($sWindowsFullPathToImageFile)
	If @error Then
		Return
	EndIf

	Local $sTranslateOutText = _SendTo_Translate_Service($sOCR_OutText)
	If @error Then
		Return
	EndIf

	If $bPutResultToClipboard Then
		ClipPut($sTranslateOutText)
	EndIf

	If $bDisplayResultToResultTab Then
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
		GUICtrlSetState($hTabResult, $GUI_SHOW)
		$iLastTab = 0
		WinActivate($hGUI_App)
		;-------------------------->>>	set focus?
	EndIf
EndFunc	;==>_OCR_and_Translate_Tray_Function
;-----------------------------------------------------------------------OCR IN TRAY MENU
Func _Only_OCR_Tray_Function()
	Local $aiPos = _ScreenRegion_GetRect(0xFF0078D7, 0x460066CC, 1)
	If @error Then
		Return
	EndIf
	Local $hCapturedBitmap = _ScreenCapture_Capture("", $aiPos[0], $aiPos[1], $aiPos[2], $aiPos[3], False)
	Local $sWindowsFullPathToImageFile = _SaveCapturedBitmap($hCapturedBitmap)
	If @error Then
		Return
	EndIf

	Local $sOCR_OutText = _SendTo_OCR_Service($sWindowsFullPathToImageFile)
	If @error Then
		Return
	EndIf

	If $bPutResultToClipboard Then
		ClipPut($sOCR_OutText)
	EndIf

	If $bDisplayResultToResultTab Then
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
		GUICtrlSetState($hTabResult, $GUI_SHOW)
		$iLastTab = 0
		WinActivate($hGUI_App)
		;-------------------------->>>	set focus?
	EndIf
EndFunc	;==>_Only_OCR_Tray_Function
;-----------------------------------------------------------------------RESULT IN TRAY MENU
Func _Result_Tray_Function()
	Local $iState = WinGetState($hGUI_App)
	If BitAND($iState, $WIN_STATE_MINIMIZED) Then
		WinSetState($hGUI_App, "", @SW_RESTORE)
	EndIf
	GUISetState(@SW_SHOW, $hGUI_App)
	If $iLastTab = 3 Then
		ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
		ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
	EndIf
	GUICtrlSetState($hTabResult, $GUI_SHOW)
	$iLastTab = 0
	WinActivate($hGUI_App)
EndFunc	;==>_Result_Tray_Function
;-----------------------------------------------------------------------CLEANING
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
;----------------------------------------------------------------------EXIT
Func _Exit_Tray_Function()
	_Cleaning()
	Exit
EndFunc	;==>_Exit_Tray_Function
;----------------------------------------------------------------------GETTING THE CAPTURE AREA
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

	While 1  And Sleep(10)
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
;----------------------------------------------------------------------CORRECTING CAPTURE DIRECTION
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
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eError]), _GetStrRes($asArrayMsg_Image[$eImageTooLarge]))				;"Error", "Image is too large, try reducing capture area."
			SetError(1)
			Return
		Else
			$sImageForOCR_Patch = $sJpg_Path
		EndIf
	Else
		$sImageForOCR_Patch = $sPng_Path
		_WinAPI_DeleteObject($hCapturedBitmap)
	EndIf
	$sWindowsFullPathToImageFile = _PathFull($sImageForOCR_Patch)
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
		Local $iState = WinGetState($hGUI_App)
		If BitAND($iState, $WIN_STATE_MINIMIZED) Then
			WinSetState($hGUI_App, "", @SW_RESTORE)
		EndIf
		GUISetState(@SW_SHOW, $hGUI_App)
		If $iLastTab = 3 Then
			ControlHide($hGUI_App, "", $hInput_HK_HotKeyOCR)
			ControlHide($hGUI_App, "", $hInput_HK_HotKeyTrans)
		EndIf
		GUICtrlSetState($hTabAPIkey, $GUI_SHOW)
		$iLastTab = 2
		SetError(31)
		Return
	EndIf
	;-----------------------------------------------------------Sending the image to OCR
	$sUnixFullPathToImageFile = StringReplace($sWindowsFullPathToImageFile, "\", "/")
	$sCurlCmd = 'curl -H "apikey:' & $sOCRAPIKey & '" --ssl-no-revoke --form "file=@' & $sUnixFullPathToImageFile & '" --form "OCREngine=2" --form "language=auto" --form "scale=true" --form "detectOrientation=true" https://api.ocr.space/Parse/Image'
	$iPID = Run($sCurlCmd, "", "", BitOR($STDERR_CHILD, $STDOUT_CHILD))
	;----------------------------------------------------------------------
	Local $hTimer = TimerInit()
	Local $fDiff
	Local $sStdErr = ""
	While 1 And Sleep(20)
		$sStdErr &= StderrRead($iPID)
		If @error Then
			ExitLoop
		EndIf
		$fDiff = TimerDiff($hTimer)
		If $fDiff > 30000 Then
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eTimeoutError]), _GetStrRes($asArrayMsg_OCR[$eOCRTimeout]))			;"Timeout Error"	"The OCR provider did not respond within 30 seconds"
			SetError(2)
			Return
		EndIf
	WEnd
	;-----------------------------------------------------------------curl Error
	Local $iCurlErrorDetect, $asCurlError

	$iCurlErrorDetect = StringRegExp($sStdErr, $sRegexCurlError)
	If $iCurlErrorDetect = 1 Then
		$asCurlError = StringRegExp($sStdErr, $sRegexCurlError, $STR_REGEXPARRAYMATCH)
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_curl[$e_curl_OCR]), $asCurlError[0])											;"curl Error during OCR"
		SetError(3)
		Return
	EndIf
	;---------------------------------------------------------------------------
	Local $sStdOut = ""

	While 1 And Sleep(20)
		$sStdOut &= StdoutRead($iPID)
		If @error Then
			ExitLoop
		EndIf

	WEnd
	;-----------------------------------------------------------------OCR Error
	Local Const $sRegexOCRIsErroredOnProcessing = '(?:\"IsErroredOnProcessing\":)(true)'
	Local Const $sRegexOCRErrorMessage = '(?:\"ErrorMessage\":\[\")(.*)(?:\"\])'
	Local Const $sRegexOCRParsedText = '(?:\"ParsedText\":\")(.*)(?:\",\"ErrorMessage\")'
	Local $iOCR_ErrorDetect, $iOCRErrorMessageDetect, $asOCRErrorMessage, $iOCRParsedTextDetect, $asOCRParsedText
	$iOCR_ErrorDetect = StringRegExp($sStdOut, $sRegexOCRIsErroredOnProcessing)
	If $iOCR_ErrorDetect = 1 Then
		$iOCRErrorMessageDetect = StringRegExp($sStdOut, $sRegexOCRErrorMessage)
		If $iOCRErrorMessageDetect = 1 Then
			$asOCRErrorMessage = StringRegExp($sStdOut, $sRegexOCRErrorMessage, $STR_REGEXPARRAYMATCH)
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_OCR[$eOCRError]), $asOCRErrorMessage[0])
			SetError(4)
			Return
		Else
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_OCR[$eOCRParsingError]), _GetStrRes($asArrayMsg_Std[$eUnknown]))			;"OCR parsing error"	"Unknown error"
			SetError(5)
			Return
		EndIf
	Else;------------------------------------------------------------OCR Result
		$iOCRParsedTextDetect = StringRegExp($sStdOut, $sRegexOCRParsedText)
		If $iOCRParsedTextDetect = 1 Then
			$asOCRParsedText = StringRegExp($sStdOut, $sRegexOCRParsedText, $STR_REGEXPARRAYMATCH)
			If $asOCRParsedText[0] == "" Then
				MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_OCR[$eOCRError]), _GetStrRes($asArrayMsg_OCR[$eOCREmpty]))				;"OCR error"	"OCR result is empty"
				SetError(6)
				Return
			EndIf
		Else
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_OCR[$eOCRParsingError]), _GetStrRes($asArrayMsg_Std[$eTextNotFound]))		;"OCR parsing error"	"Text not found"
			SetError(7)
			Return
		EndIf
	EndIf

	Local $sDecodeParsedTextTo_UTF8 = _WinAPI_MultiByteToWideChar($asOCRParsedText[0], 65001, 0, True)
	Const $sRegexLineBreak = '(?m)(?<=[\.|\?|\!])(\\n)'
	Local $sOCR_OutText =  StringRegExpReplace($sDecodeParsedTextTo_UTF8, $sRegexLineBreak, @CRLF)
	$sOCR_OutText = StringReplace($sOCR_OutText, '\n', ' ')
	$sOCR_OutText = StringReplace($sOCR_OutText, '\"', '"')
	$sOCR_OutText = StringReplace($sOCR_OutText, '\\', '\')
	$sOCR_OutText = StringReplace($sOCR_OutText, '\/', '/')   ;\b,\r,\f,\t

	Return $sOCR_OutText
EndFunc ;==>_SendTo_OCR_Service
;----------------------------------------------------------------------SENDING TO TRANSLATE SERVICE
Func _SendTo_Translate_Service($sOCR_OutText)

	Local $sURI_TextForTranslate = _URIEncode($sOCR_OutText)
	Local $sCurlCmd = 'curl --ssl-no-revoke -X GET "https://www.googleapis.com/language/translate/v2?key=' & $sTranslateAPIKey & '&target=' & $sTargetLanguage & '&format=text&q=' & $sURI_TextForTranslate & '"'
	Local $iPID = Run($sCurlCmd, "", "", BitOR($STDERR_CHILD, $STDOUT_CHILD))
	;----------------------------------------------------------------------------------
	Local $hTimer = TimerInit()
	Local $sStdErr = ""
	While 1 And Sleep(20)
		$sStdErr &= StderrRead($iPID)
		If @error Then
			ExitLoop
		EndIf
		$fDiff = TimerDiff($hTimer)
		If $fDiff > 30000 Then
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Std[$eTimeoutError]), _GetStrRes($asArrayMsg_Translate[$eTransTimeout]))	;"Timeout Error"	"The Translation provider did not respond within 30 seconds."
			SetError(8)
			Return
		EndIf
	WEnd
	;----------------------------------------------------------------------curl Error
	Local $iCurlErrorDetect, $asCurlError
	$iCurlErrorDetect = StringRegExp($sStdErr, $sRegexCurlError)
	If $iCurlErrorDetect = 1 Then
		$asCurlError = StringRegExp($sStdErr, $sRegexCurlError, $STR_REGEXPARRAYMATCH)
		MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_curl[$e_curl_Translate]), $asCurlError[0])										;"curl Error during translation"
		SetError(9)
		Return
	EndIf
	;--------------------------------------------------------------------------------
	Local $sStdOut = ""

	While 1 And Sleep(20)
		$sStdOut &= StdoutRead($iPID)
		If @error Then
			ExitLoop
		EndIf

	WEnd
	;---------------------------------------------------------------Translation Error
	Const $sRegexTranslateErrorResponse = '(^\{\s*\"error\":)'
	Const $sRegexTranslateErrorCode = '(?:\"code\":\s)(\d*)'
	Const $sRegexTranslateErrorMessage = '(?:\"message\":\s\")([^\"]*)'
	Const $sRegexTranslateParsedText = '(?:\"translatedText\":\s\")(.*)(?:\",\s*\"detectedSourceLanguage\")'
	Local $iTranslateErrorDetect, $iTranslateErrorCodeDetect, $iTranslateErrorMessageDetect, $asTranslateErrorCode, $asTranslateErrorMessage, $iTranslateParsedTextDetect, $asTranslateParsedText
	$iTranslateErrorDetect = StringRegExp($sStdOut, $sRegexTranslateErrorResponse)
	If $iTranslateErrorDetect = 1 Then
		$iTranslateErrorCodeDetect = StringRegExp($sStdOut, $sRegexTranslateErrorCode)
		$iTranslateErrorMessageDetect = StringRegExp($sStdOut, $sRegexTranslateErrorMessage)
		If $iTranslateErrorCodeDetect = 1 And $iTranslateErrorMessageDetect = 1 Then
			$asTranslateErrorCode = StringRegExp($sStdOut, $sRegexTranslateErrorCode, $STR_REGEXPARRAYMATCH)
			$asTranslateErrorMessage = StringRegExp($sStdOut, $sRegexTranslateErrorMessage, $STR_REGEXPARRAYMATCH)
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Translate[$eTransError]), $asTranslateErrorCode[0] & ": " & $asTranslateErrorMessage[0])	;"Translation error"
			SetError(10)
			Return
		Else
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Translate[$eTransParsingError]), _GetStrRes($asArrayMsg_Std[$eUnknown]))					;"Translation parsing error" 	"Unknown error"
			SetError(11)
			Return
		EndIf
	Else	;--------------------------------------------------------Translation Result
		$iTranslateParsedTextDetect = StringRegExp($sStdOut, $sRegexTranslateParsedText)
		If $iTranslateParsedTextDetect = 1 Then
			$asTranslateParsedText = StringRegExp($sStdOut, $sRegexTranslateParsedText, $STR_REGEXPARRAYMATCH)
			If $asTranslateParsedText[0] == "" Then
				MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Translate[$eTransError]), _GetStrRes($asArrayMsg_Translate[$eTransEmpty]))				;"Translation error"	"Translation result is empty"
				SetError(12)
				Return
			EndIf
		Else
			MsgBox($MB_ICONERROR, _GetStrRes($asArrayMsg_Translate[$eTransParsingError]), _GetStrRes($asArrayMsg_Std[$eTextNotFound]))				;"Translation parsing error"	"Text not found"
			SetError(13)
			Return
		EndIf
	EndIf

	Local $sDecodeTranslateParsedTextTo_UTF8 = _WinAPI_MultiByteToWideChar($asTranslateParsedText[0], 65001, 0, True)
	Local $sTranslateOutText = StringRegExpReplace($sDecodeTranslateParsedTextTo_UTF8, '(\\r\\n)', @CRLF)
	$sTranslateOutText = StringReplace($sTranslateOutText, '\n', @CRLF)
	$sTranslateOutText = StringReplace($sTranslateOutText, '\"', '"')
	$sTranslateOutText = StringReplace($sTranslateOutText, '\\', '\')
	Return $sTranslateOutText
EndFunc ;==>_SendTo_Translate_Service
;-----------------------------------------------------------------------DETECTING DARK-LIGHT THEME
Func _Detect_Dark_Light_Theme()
	Local $sArchitecture = @OSArch
	If $sArchitecture = "X64" Then
		$nAppsUseLightTheme = RegRead("HKCU64\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme")
		$nSystemUsesLightTheme = RegRead("HKCU64\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme")
	ElseIf $sArchitecture = "X86" Then
		$nAppsUseLightTheme = RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme")
		$nSystemUsesLightTheme = RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme")
	Else
		MsgBox($MB_ICONINFORMATION, "IA64", "Itanium is not supported")
		Exit
	EndIf
EndFunc
;-----------------------------------------------------------------------ACTIVITY OF CONTEXT MENU ITEMS FOR INPUT
Func _ActivityContextMenuItem_Input()
	ClipGet()
	If @error Then
		If _GUICtrlMenu_GetItemEnabled($hContextMenu_Input, 2) Then
			_GUICtrlMenu_SetItemDisabled($hContextMenu_Input, 2)	;Paste
		EndIf
	Else
		If _GUICtrlMenu_GetItemDisabled($hContextMenu_Input, 2) Then
			_GUICtrlMenu_SetItemEnabled($hContextMenu_Input, 2)		;Paste
		EndIf
	EndIf
	;-------------------------------------------------
	$aSel = _GUICtrlEdit_GetSel($idInput_API_OCRAPIKey)
	$iSelect = $aSel[1] - $aSel[0]
	If $iSelect = 0 Then
		If _GUICtrlMenu_GetItemEnabled($hContextMenu_Input, 0) Then
			_GUICtrlMenu_SetItemDisabled($hContextMenu_Input, 0)	;Cut
		EndIf
		If _GUICtrlMenu_GetItemEnabled($hContextMenu_Input, 1) Then
			_GUICtrlMenu_SetItemDisabled($hContextMenu_Input, 1)	;Copy
		EndIf
		If _GUICtrlMenu_GetItemEnabled($hContextMenu_Input, 3) Then
			_GUICtrlMenu_SetItemDisabled($hContextMenu_Input, 3)	;Delete
		EndIf
	Else
		If _GUICtrlMenu_GetItemDisabled($hContextMenu_Input, 0) Then
			_GUICtrlMenu_SetItemEnabled($hContextMenu_Input, 0)		;Cut
		EndIf
		If _GUICtrlMenu_GetItemDisabled($hContextMenu_Input, 1) Then
			_GUICtrlMenu_SetItemEnabled($hContextMenu_Input, 1)		;Copy
		EndIf
		If _GUICtrlMenu_GetItemDisabled($hContextMenu_Input, 3) Then
			_GUICtrlMenu_SetItemEnabled($hContextMenu_Input, 3)		;Delete
		EndIf
	EndIf
	;--------------------------------------------------
	$iAll = StringLen(GUICtrlRead($idInput_API_OCRAPIKey))
	If $iAll = 0 Or $iAll = $iSelect Then
		If _GUICtrlMenu_GetItemEnabled($hContextMenu_Input, 5) Then
			_GUICtrlMenu_SetItemDisabled($hContextMenu_Input, 5)	;Select All
		EndIf
	Else
		If _GUICtrlMenu_GetItemDisabled($hContextMenu_Input, 5) Then
			_GUICtrlMenu_SetItemEnabled($hContextMenu_Input, 5)		;Select All
		EndIf
	EndIf
EndFunc
;----------------------------------------------------------------------ACTIVITY OF CONTEXT MENU ITEMS FOR EDIT
Func _ActivityContextMenuItem_Edit()
	$aSel = _GUICtrlEdit_GetSel($idEdit_RSLT)
	$iSelect = $aSel[1] - $aSel[0]
	If $iSelect = 0 Then
		If _GUICtrlMenu_GetItemEnabled($hContextMenu_Edit, 0) Then
			_GUICtrlMenu_SetItemDisabled($hContextMenu_Edit, 0)		;Copy
		EndIf
	Else
		If _GUICtrlMenu_GetItemDisabled($hContextMenu_Edit, 0) Then
			_GUICtrlMenu_SetItemEnabled($hContextMenu_Edit, 0)		;Copy
		EndIf
	EndIf
	;--------------------------------------------------
	$iAll = StringLen(GUICtrlRead($idEdit_RSLT))
	If $iAll = 0 Or $iAll = $iSelect Then
		If _GUICtrlMenu_GetItemEnabled($hContextMenu_Edit, 2) Then
			_GUICtrlMenu_SetItemDisabled($hContextMenu_Edit, 2)		;Select All
		EndIf
	Else
		If _GUICtrlMenu_GetItemDisabled($hContextMenu_Edit, 2) Then
			_GUICtrlMenu_SetItemEnabled($hContextMenu_Edit, 2)		;Select All
		EndIf
	EndIf
EndFunc
