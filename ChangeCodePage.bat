@ECHO OFF
TITLE Change Codepage for ArcGIS
:: --------------------------------------------------------------
::  ChangeCodePage.bat
:: --------------------------------------------------------------
::  Updates setting for dbfDefault (Shapefiles Codepage)
::  for ArcGIS Desktop 10.x or / and ArcGIS Pro
:: --------------------------------------------------------------
::  Copyright 2025 - Dr. Panagiotis E. Papazoglou
::  Use it at your own risk.
:: --------------------------------------------------------------
::  ������餜� �� �矣��� dbfDefault (��������垩� Shapefiles)
::  ��� ArcGIS Desktop 10.x � / ��� ArcGIS Pro
:: --------------------------------------------------------------
::  ���������� �����飘�� 2025 - ��. ������髞� �. ���ᝦ���� 
::  ����������㩫� �� �� ���� ��� ���礞.
:: --------------------------------------------------------------
SETLOCAL EnableDelayedExpansion
FOR /F "tokens=3" %%L IN (
  'reg query "HKCU\Control Panel\International" /v LocaleName ^| find "REG_SZ"'
) DO SET "LOC=%%L"
IF /I "%LOC:~0,2%"=="el" (
	SET "LANGMODE=GR"
) ELSE (
	SET "LANGMODE=EN"
)
IF "%LANGMODE%"=="GR" (
	SET "MSG_TITLE1=   ������ ����������⤞� ��������垩�� ��� �� Shapefiles"
	SET "MSG_TITLE2=        ��� ArcGIS Desktop 10.x � / ��� ArcGIS Pro"
	SET "MSG_APPSEL_TITLE= ����⥫� �������� ��� ��� ���� �� �������� � ��������垩�"
	SET "MSG_APPSEL_OPT="
	SET "MSG_APPSEL_EXIT=륦���"
	SET "MSG_APPSEL_INPUT=����������㩫� 1, 2, 3 � 4: "
	SET "MSG_CPTITLE=          �������� ArcGIS Desktop ��� ArcGIS Pro"
	SET "MSG_CPTITLE1=                   �������� ArcGIS %%TARGET%%"
	SET "MSG_CPSEL_TITLE=����⥫� ��������垩�:"
	SET "MSG_CPSEL_TITLE2=    ����������㩫� ��� ������ ������� ��� ���㩫� ENTER"
	SET "CPDESC1= 1)  437 - �.�.�. (DOS)                25)   1252 - ������ ���駞� (Windows)"
	SET "CPDESC2= 2)  708 - ������� (ASMO 708)          26)   1253 - �������� (Windows ANSI)"
	SET "CPDESC3= 3)  720 - ������� (������� ASMO)      27)   1254 - �������� (Windows)"
	SET "CPDESC4= 4)  737 - �������� (DOS)              28)   1255 - ������ (Windows)"
	SET "CPDESC5= 5)  775 - ������� (DOS)               29)   1256 - ������� (Windows)"
	SET "CPDESC6= 6)  850 - ������ ���駞� (DOS)       30)   1257 - ������� (Windows)"
	SET "CPDESC7= 7)  852 - �������� ���駞� (DOS)     31)   1258 - ���������� (Windows)"
	SET "CPDESC8= 8)  855 - ��������� (DOS)             32)   Big5 - Big-5 (����������� ��� 950)"
	SET "CPDESC9= 9)  857 - �������� (DOS)              33)   SJIS - SJIS (����������� ��� 932)"
	SET "CPDESC10=10)  860 - ����������� (DOS)           34)  88591 - ��������-1 (������ �.�.)"
	SET "CPDESC11=11)  861 - ��������� (DOS)             35)  88592 - ��������-2 (��������/��������� �.�.)"
	SET "CPDESC12=12)  862 - ������ (DOS)               36)  88593 - ��������-3 (�櫠� �.�.)"
	SET "CPDESC13=13)  863 - ������� �������� (DOS)      37)  88594 - ��������-4 (�樜�� �.�.)"
	SET "CPDESC14=14)  864 - ������� (864)               38)  88595 - ���������"
	SET "CPDESC15=15)  865 - ��������� (DOS)             39)  88596 - �������"
	SET "CPDESC16=16)  866 - �੠�� (��������� DOS)      40)  88597 - ��������"
	SET "CPDESC17=17)  869 - ����⨤� �������� (DOS)     41)  88598 - ������"
	SET "CPDESC18=18)  932 - ���ठ�� (Shift-JIS)        42)  88599 - ��������-5 (��������)"
	SET "CPDESC19=19)  936 - ���❠�� (PRC GBK)          43) 885910 - ��������-6 (���������)"
	SET "CPDESC20=20)  949 - ����᫠�� (UHC)             44) 885911 - ��䢘�����"
	SET "CPDESC21=21)  950 - ���❠�� (Big-5 - ����)   45) 885913 - ����������"
	SET "CPDESC22=22)  448 - ALARABI (��坜� 448)        46) 885915 - ��������-9 (�����ਫ਼ ������ �.�.)"
	SET "CPDESC23=23) 1250 - �������� ���駞� (Windows) 47)  UTF-8 - Unicode (65001)"
	SET "CPDESC24=24) 1251 - ��������� (Windows)         48)   UTF8 - Unicode (����������� ��� 65001)"
	SET "MSG_CPSEL_PROMPT=������ ������� (�.�. 26 ��� ��������): "
	SET "MSG_PRESS_ANY_KEY=���㩫� ������㧦�� ��㡫�� ��� �� ���᚜�� ⤘� ������ ��� �� �婫�!"
	SET "MSG_SUMMARY_TITLE=                          �礦��:"
	SET "MSG_SUMMARY_BOTH=   �� �������� ArcGIS �� ������������ ��� �� shapefiles"
	SET "MSG_SUMMARY_SINGLE= � �������� ArcGIS %%TARGET%% �� ������������ ��� �� shapefiles"
	SET "MSG_SUMMARY_CODEPAGE=           ��� ��������垩� '%%CodePage%%' � ����������"
	SET "MSG_CONFIRM_PROMPT=�⢜�� �� �����橜�� ��� ������� ���; (���/): "
	SET "MSG_CANCEL=� ��������� ����韞�� ��� ��� ��㩫�."
	SET "MSG_INVALID_CONFIRM=�� ⚡��� ��ᤫ���. �����ᩫ� ����."
	SET "MSG_INVALID_APPSEL=�� ⚡��� �������. ����������."
	SET "MSG_INVALID_CODEPAGE=�� ⚡��� ������� ���������囘�. ����⥫� ��� �� �婫�."
	SET "MSG_SUCCESS_TITLE1=              � �������� ��� �☪ ��������垩��"
	SET "MSG_SUCCESS_TITLE2=           ��� �� shapefiles �������韞�� �������!"
	SET "MSG_CHANGE_CP_FOR_D========= ������ ��������垩�� ��� �� ArcGIS Desktop ========"
	SET "MSG_CHANGE_CP_FOR_P=========== ������ ��������垩�� ��� �� ArcGIS Pro =========="
	SET "MSG_KEEP_PROMPT=�⢜�� �� ������㩜�� �� ���៬�� �������;"
	SET "MSG_KEEP_YES=���"
	SET "MSG_KEEP_��="
	SET "MSG_KEEP_CHOOSE=����⥫� 1 � 2: "
) ELSE (
	SET "MSG_TITLE1=           Change default codepage of Shapefiles"
	SET "MSG_TITLE2=        for ArcGIS Desktop 10.x or / and ArcGIS Pro"
	SET "MSG_APPSEL_TITLE=Select the application whose shapefile encoding will change"
	SET "MSG_APPSEL_OPT=Both"
	SET "MSG_APPSEL_EXIT=Exit"
	SET "MSG_APPSEL_INPUT=Type 1, 2, 3 or 4: "
	SET "MSG_CPTITLE=               ArcGIS Desktop and ArcGIS Pro"
	SET "MSG_CPTITLE1=                        ArcGIS %%TARGET%%"
	SET "MSG_CPSEL_TITLE=Select code page:"
	SET "MSG_CPSEL_TITLE2=         Type the number you choose and press ENTER"
	SET "CPDESC1= 1)  437 - USA (DOS)                 25)   1252 - Western Europe (Windows)"
	SET "CPDESC2= 2)  708 - Arabic (ASMO 708)         26)   1253 - Greek (Windows ANSI)"
	SET "CPDESC3= 3)  720 - Arabic (Transparent ASMO) 27)   1254 - Turkish (Windows)"
	SET "CPDESC4= 4)  737 - Greek (DOS)               28)   1255 - Hebrew (Windows)"
	SET "CPDESC5= 5)  775 - Baltic (DOS)              29)   1256 - Arabic (Windows)"
	SET "CPDESC6= 6)  850 - Western Europe (DOS)      30)   1257 - Baltic (Windows)"
	SET "CPDESC7= 7)  852 - Central Europe (DOS)      31)   1258 - Vietnamese (Windows)"
	SET "CPDESC8= 8)  855 - Cyrillic (DOS)            32)   Big5 - Big-5 (alternative to 950)"
	SET "CPDESC9= 9)  857 - Turkish (DOS)             33)   SJIS - SJIS (alternative to 932)"
	SET "CPDESC10=10)  860 - Portuguese (DOS)          34)  88591 - Latin-1 (Western Europe)"
	SET "CPDESC11=11)  861 - Icelandic (DOS)           35)  88592 - Latin-2 (Central/Eastern Europe)"
	SET "CPDESC12=12)  862 - Hebrew (DOS)              36)  88593 - Latin-3 (Southern Europe)"
	SET "CPDESC13=13)  863 - French Canadian (DOS)     37)  88594 - Latin-4 (Northern Europe)"
	SET "CPDESC14=14)  864 - Arabic (864)              38)  88595 - Cyrillic"
	SET "CPDESC15=15)  865 - Norwegian (DOS)           39)  88596 - Arabic"
	SET "CPDESC16=16)  866 - Russian (Cyrillic DOS)    40)  88597 - Greek"
	SET "CPDESC17=17)  869 - Modern Greek (DOS)        41)  88598 - Hebrew"
	SET "CPDESC18=18)  932 - Japanese (Shift-JIS)      42)  88599 - Latin-5 (Turkish)"
	SET "CPDESC19=19)  936 - Chinese (PRC GBK)         43) 885910 - Latin-6 (Norwegian)"
	SET "CPDESC20=20)  949 - Korean (UHC)              44) 885911 - Thai"
	SET "CPDESC21=21)  950 - Chinese (Big-5 - Taiwan)  45) 885913 - Lithuanian"
	SET "CPDESC22=22)  448 - ALARABI (defines 448)     46) 885915 - Latin-9 (Western Europe update)"
	SET "CPDESC23=23) 1250 - Central Europe (Windows)  47)  UTF-8 - Unicode (65001)"
	SET "CPDESC24=24) 1251 - Cyrillic (Windows)        48)   UTF8 - Unicode (alternative to 65001)"
	SET "MSG_CPSEL_PROMPT=Selection number (e.g. 26 for Greek): "
	SET "MSG_PRESS_ANY_KEY=Press any key in order to insert a number from the above list!"
	SET "MSG_SUMMARY_TITLE=Summary:"
	SET "MSG_SUMMARY_BOTH=      The ArcGIS applications will use for shapefiles"
	SET "MSG_SUMMARY_SINGLE=    The application ArcGIS %%TARGET%% will use for shapefiles"
	SET "MSG_SUMMARY_CODEPAGE=              the code page '%%CodePage%%' as default"
	SET "MSG_CONFIRM_PROMPT=Do you want to apply your selection? (Yes/No): "
	SET "MSG_CANCEL=The process was cancelled by the user."
	SET "MSG_INVALID_CONFIRM=Invalid answer. Please try again."
	SET "MSG_INVALID_APPSEL=Invalid selection. Exiting."
	SET "MSG_INVALID_CODEPAGE=Invalid code page selection. Select from the list."
	SET "MSG_SUCCESS_TITLE1=                 The new shapefile encoding"
	SET "MSG_SUCCESS_TITLE2=               has been applied successfully!"
	SET "MSG_CHANGE_CP_FOR_D============= Change encoding for ArcGIS Desktop ============"
	SET "MSG_CHANGE_CP_FOR_P=============== Change encoding for ArcGIS Pro =============="
	SET "MSG_KEEP_PROMPT=Do you want to keep the window open?"
	SET "MSG_KEEP_YES=Yes"
	SET "MSG_KEEP_��=No"
	SET "MSG_KEEP_CHOOSE=Choose 1 or 2: "
)
CLS
ECHO.
ECHO ============================================================
ECHO %MSG_TITLE1%
ECHO %MSG_TITLE2%
ECHO ============================================================
ECHO %MSG_APPSEL_TITLE%
ECHO ============================================================
ECHO   1) ArcGIS Desktop
ECHO   2) ArcGIS Pro
ECHO   3) %MSG_APPSEL_OPT%
ECHO   4) %MSG_APPSEL_EXIT%
ECHO.
CHOICE /C 1234 /N /M "%MSG_APPSEL_INPUT%"
IF ERRORLEVEL 4 GOTO :EOF
IF ERRORLEVEL 3 (SET "TARGET=Both" & GOTO SELECT_CP)
IF ERRORLEVEL 2 (SET "TARGET=Pro" & GOTO SELECT_CP)
IF ERRORLEVEL 1 (SET "TARGET=Desktop" & GOTO SELECT_CP)
ECHO %MSG_INVALID_APPSEL%
GOTO :EOF

:SELECT_CP
CLS
ECHO.
ECHO ============================================================
IF /I "%TARGET%"=="Both" (
	ECHO %MSG_CPTITLE%
) ELSE (
	CALL ECHO %MSG_CPTITLE1%
)
ECHO ============================================================
ECHO %MSG_CPSEL_TITLE%
ECHO %MSG_CPSEL_TITLE2%
ECHO ============================================================
ECHO %CPDESC1%
ECHO %CPDESC2%
ECHO %CPDESC3%
ECHO %CPDESC4%
ECHO %CPDESC5%
ECHO %CPDESC6%
ECHO %CPDESC7%
ECHO %CPDESC8%
ECHO %CPDESC9%
ECHO %CPDESC10%
ECHO %CPDESC11%
ECHO %CPDESC12%
ECHO %CPDESC13%
ECHO %CPDESC14%
ECHO %CPDESC15%
ECHO %CPDESC16%
ECHO %CPDESC17%
ECHO %CPDESC18%
ECHO %CPDESC19%
ECHO %CPDESC20%
ECHO %CPDESC21%
ECHO %CPDESC22%
ECHO %CPDESC23%
ECHO %CPDESC24%
ECHO.
SET "CPSel="
SET /P CPSel="%MSG_CPSEL_PROMPT%"
SET "CodePage="
FOR %%# IN (
	"1=437"		"2=708"		"3=720"		"4=737"
	"5=775"		"6=850"		"7=852"		"8=855"
	"9=857"		"10=860"	"11=861"	"12=862"
	"13=863"	"14=864"	"15=865"	"16=866"
	"17=869"	"18=932"	"19=936"	"20=949"
	"21=950"	"22=448"	"23=1250"	"24=1251"
	"25=1252"	"26=1253"	"27=1254"	"28=1255"
	"29=1256"	"30=1257"	"31=1258"	"32=Big5"
	"33=SJIS"	"34=88591"	"35=88592"	"36=88593"
	"37=88594"	"38=88595"	"39=88596"	"40=88597"
	"41=88598"	"42=88599"	"43=885910"	"44=885911"
	"45=885913"	"46=885915"	"47=UTF-8"	"48=UTF8"
) DO (
	FOR /F "tokens=1,2 delims==" %%A IN ("%%~#") DO (
		IF "!CPSel!"=="%%A" SET "CodePage=%%B"
	)
)
IF NOT DEFINED CodePage (
	ECHO %MSG_INVALID_CODEPAGE%
	SET /P ="%MSG_PRESS_ANY_KEY%" < NUL
	PAUSE > NUL
	GOTO SELECT_CP
)
CLS
ECHO.
ECHO ============================================================
ECHO %MSG_SUMMARY_TITLE%
IF /I "%TARGET%"=="Both" (
	ECHO %MSG_SUMMARY_BOTH%
) ELSE (
	CALL ECHO %MSG_SUMMARY_SINGLE%
)
CALL ECHO %MSG_SUMMARY_CODEPAGE%
ECHO ============================================================

:CONFIRM
ECHO.
SET /P Confirm="%MSG_CONFIRM_PROMPT%"
IF /I "%Confirm%"=="Y" GOTO CONTINUE
IF /I "%Confirm%"=="N" GOTO CANCEL
IF /I "%Confirm%"=="�" GOTO CONTINUE
IF /I "%Confirm%"=="�" GOTO CANCEL
IF /I "%Confirm%"=="�" GOTO CANCEL
ECHO %MSG_INVALID_CONFIRM%
GOTO CONFIRM

:CANCEL
CLS
ECHO %MSG_CANCEL%
cmd /k

:CONTINUE
IF /I "%TARGET%"=="Desktop" GOTO APPLY_DESKTOP
IF /I "%TARGET%"=="Pro" GOTO APPLY_PRO
IF /I "%TARGET%"=="Both" GOTO APPLY_DESKTOP
GOTO :EOF

:APPLY_DESKTOP
ECHO.
ECHO %MSG_CHANGE_CP_FOR_D%
FOR %%N IN (0 1 2 3 4 5 6 7 8) DO (
	REG QUERY "HKCU\Software\ESRI\Desktop10.%%N" /ve >NUL 2>&1 && (
		REG ADD "HKCU\Software\ESRI\Desktop10.%%N\Common\CodePage" ^
			/v dbfDefault /t REG_SZ /d "!CodePage!" /f >NUL
	)
)
IF /I NOT "%TARGET%"=="Both" GOTO SUCCESS

:APPLY_PRO
ECHO.
ECHO %MSG_CHANGE_CP_FOR_P%
REG QUERY "HKCU\Software\ESRI\ArcGISPro" /ve >NUL 2>&1 && (
	REG ADD "HKCU\Software\ESRI\ArcGISPro\Common\CodePage" ^
		/v dbfDefault /t REG_SZ /d "!CodePage!" /f >NUL
)
GOTO SUCCESS

:SUCCESS
ECHO.
ECHO ============================================================
ECHO %MSG_SUCCESS_TITLE1%
ECHO %MSG_SUCCESS_TITLE2%
ECHO ============================================================
ECHO %MSG_KEEP_PROMPT%
ECHO 1) %MSG_KEEP_YES%
ECHO 2) %MSG_KEEP_��%
ECHO.
CHOICE /C 12 /N /M "%MSG_KEEP_CHOOSE%"
IF ERRORLEVEL 2 EXIT
IF ERRORLEVEL 1 (
	CLS
	cmd /k
)