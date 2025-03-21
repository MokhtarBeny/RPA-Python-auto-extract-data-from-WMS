;---------------------------------------------------------------------------------------------------------------
; Fichier extraction des données de SAP en temps réel
; transaction VL06I de P48
;
;
;
;
;

Global $LIVRAISON, $REFERENCE,$MAGASIN,$ID_TRANSPORT,$STATUT_ENTREE_STOCK,$STATUT_GLOBAL_MVMT_STOCK,$FOURNISSEUR_P48,$CREE_LE,$LIVRAISON_EXTERNE


#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <_SQL.au3>
#include <Date.au3>

_AU3RecordSetup()

Global $Nom_Table='SCI_BASE_STANDARDS'

Local $Date = @MDAY & "-" & @MON & "-" & @YEAR

Global $Chemin_Fichier = "U:\Flux\Centrale\SCI\Exportation\"
Global $Nom_fichier_old = "FR50L.STK.006_ Ref STD et emplacement.csv"
Global $Nom_fichier_new = "FR50L.STK.006_ Ref STD et emplacement_MODIFIED.csv"

Global $NOM_REPERTOIRE_SERVEUR = 'autoit$'

Global $DefaultInstance = 'FRROYESQL01P'
Global $DefaultBDD= 'BASE_PBI'

$INSTANCE_SQL2  = 'FRROYEAPP2'
$NOM_BDD2	= 'CLIRDATA'


;on a une seule base
Global $NOMBRE_BDD = 1


If $cmdline[0] >= 1 Then
		Switch $cmdline[1]
		   ; cas par défault dans le script
			case 0
				$INSTANCE_SQL    = $DefaultInstance
				$NOM_BDD         = $DefaultBDD
			;cas qui utilise les arguments du planificateur
			case 1
				$INSTANCE_SQL     = $cmdline[2]
				$NOM_BDD          = $cmdline[3]
			;cas qui utilise plusieurs paramètres >> planificateur et variable par défault
			case 2
				;on a 2 DB
				$NOMBRE_BDD = 2

				$INSTANCE_SQL     = $DefaultInstance
				$NOM_BDD         = $DefaultBDD

				$INSTANCE_SQL2     = $cmdline[2]
				$NOM_BDD2         = $cmdline[3]
		EndSwitch
		Else
			GLOBAL $INSTANCE_SQL     = $DefaultInstance
			GLOBAL $NOM_BDD         = $DefaultBDD
Endif

;on enlève les entêtes youpi

Local $hFileOpen, $sFileRead, $hFileOpen_out, $sFile_out
$hFileOpen = FileOpen($Chemin_Fichier & $Nom_fichier_old, $FO_READ)

If $hFileOpen = -1 Then
	MsgBox(0, "", "Une erreur est survenue lors de la lecture du fichier.",10)
	Exit
 EndIf

$hFileOpen_out = FileOpen($Chemin_Fichier & $Nom_fichier_new, $FO_OVERWRITE )

If $hFileOpen_out = -1 Then
	MsgBox(0, "", "Une erreur est survenue lors de l'ouverture du fichier de sortie.",10)
	Exit
EndIf

$ARTICLE			= ""
$EMPLACEMENT		= ""
$ALLOCATION			= ""

$sFileRead = FileReadLine($hFileOpen,2)
while (@error <> -1)
	Local $Zone_decoupee = StringSplit($sFileRead, ";")
	$ARTICLE			= StringFormat("%-8.8s",StringStripWS($Zone_decoupee[1], $STR_STRIPLEADING + $STR_STRIPTRAILING))
	$EMPLACEMENT		= StringFormat("%-10.10s",StringStripWS($Zone_decoupee[2], $STR_STRIPLEADING + $STR_STRIPTRAILING))
	$ALLOCATION 		= StringFormat("%-20.20s",StringStripWS($Zone_decoupee[3], $STR_STRIPLEADING + $STR_STRIPTRAILING))
	FileWrite($hFileOpen_out,  $ARTICLE &";"& $EMPLACEMENT &";"& $ALLOCATION & @CRLF)
	$sFileRead= FileReadLine($hFileOpen)

WEnd

FileClose($hFileOpen)
FileClose($hFileOpen_out)


; déplacement du fichier sur le serveur SQL

; Si un paramètre est passé, alors on ajoute dans l'ancienne base
if $NOMBRE_BDD=2 Then												; Si relance on doit indiquer la date dont on n'a pas les données
	FileCopy($Chemin_Fichier&$Nom_Fichier_new,"\\"&$INSTANCE_SQL2&"\"&$NOM_REPERTOIRE_SERVEUR&"\"&$Nom_Fichier_new, $FC_OVERWRITE )

	Insert_BDD ($INSTANCE_SQL2,$NOM_BDD2)

	Envoie_Mail("roye.dl.it@loreal.com","","Table "&$Nom_Table&" OK","",$Nom_Table,$INSTANCE_SQL2,$NOM_BDD2)
EndIf
; Fin ancienne base

; déplacement du fichier sur le serveur SQL
FileMove($Chemin_Fichier&$Nom_Fichier_new,"\\"&$INSTANCE_SQL&"\"&$NOM_REPERTOIRE_SERVEUR&"\"&$Nom_Fichier_new, $FC_OVERWRITE )

Insert_BDD ($INSTANCE_SQL,$NOM_BDD)


Envoie_Mail("roye.dl.it@loreal.com","","Table "&$Nom_Table&" OK","",$Nom_Table,$INSTANCE_SQL,$NOM_BDD)


;fin main()

FileDelete ($Chemin_Fichier & $Nom_fichier_old)

;----------------------------------------------------------------------------------------------------------------------
;fonction de mise a jour dans la BDD
;----------------------------------------------------------------------------------------------------------------------


func Insert_BDD ($INSTANCE_SQL,$NOM_BDD)
Local $SQL_err

_SQL_Startup()

_SQL_Connect(-1, $INSTANCE_SQL, $NOM_BDD, '', '')

if @error Then
	MsgBox(0, "ERROR", "Failed to connect to the database",10)
	_SQL_Close()
	; Ferme le handle renvoyé par FileOpen.
	return 0
EndIf



;on supprime les données

$sQuery = "SET DATEFORMAT dmy; delete from "&$Nom_Table
ConsoleWrite($sQuery & @CRLF)

$SQL_err = _SQL_Execute(-1, $sQuery)
ConsoleWrite("exec delete" & @CRLF)

if $SQL_err = $SQL_ERROR Then
	MsgBox (0, "erreur", "erreur delete "&$Nom_Table,10)
	Return 0
	_SQL_Close()
EndIf
ConsoleWrite("Pas d'erreur de delete" & @CRLF)


; avant de les réinjecter

$sQuery = "SET DATEFORMAT dmy; BULK INSERT "&$Nom_Table&" FROM '\\"&$INSTANCE_SQL&"\"&$NOM_REPERTOIRE_SERVEUR&"\"&$Nom_Fichier_new &"' WITH ( FIELDTERMINATOR = ';', ROWTERMINATOR = '\n', firstrow = 1 )"
ConsoleWrite($sQuery & @CRLF)


$SQL_err = _SQL_Execute(-1, $sQuery)
ConsoleWrite("exec insert" & @CRLF)
if $SQL_err = $SQL_ERROR Then
	MsgBox (0, "erreur", "erreur insert "&$Nom_Table,10)
	Return 0
	_SQL_Close()
EndIf



ConsoleWrite("Pas d'erreur d'insert" & @CRLF)
_SQL_Close()
Return 1


EndFunc

;----------------------------------------------------------------------------------------------------------------------
;
;----------------------------------------------------------------------------------------------------------------------

Func _Au3RecordSetup()
Opt('WinWaitDelay',1000)
Opt('WinDetectHiddenText',1)
Opt('MouseCoordMode',0)
Opt('SendKeyDelay',100)
Local $aResult = DllCall('User32.dll', 'int', 'GetKeyboardLayoutNameW', 'wstr', '')
If $aResult[1] <> '0000040C' Then
  MsgBox(64, 'Warning', 'Recording has been done under a different Keyboard layout' & @CRLF & '(0000040C->' & $aResult[1] & ')',10)
EndIf

EndFunc

;----------------------------------------------------------------------------------------------------------------------
;
;----------------------------------------------------------------------------------------------------------------------

Func _WinWaitActivate($title,$text,$timeout=200)
	WinWait($title,$text,$timeout)
	If Not WinActive($title,$text) Then WinActivate($title,$text)
	If WinWaitActive($title,$text,$timeout) = 0 Then
	   run ("C:\WINDOWS\system32\taskkill.exe /F /IM saplogon.exe")
	   ;run ("C:\WINDOWS\system32\taskkill.exe /F /IM excel.exe")
	   Exit
	EndIf
 EndFunc


;---------------------------------------------------------------------------
; fonction pour enlever les virgules
;---------------------------------------------------------------------------

Func enleve_virgule($decoup,$carac)
Local $zone = StringFormat("%-"&$carac&"."&$carac&"s",StringReplace($Zone_decoupee[$decoup], ",", ""))
return ($zone)
EndFunc


;---------------------------------------------------------------------------
; on supprime tous les caractères spéciaux ex: caractère découvert 
;---------------------------------------------------------------------------
Func Sup_Car_Sp($Z_Zone)

Local $aArray = StringToASCIIArray($Z_Zone)
local $Z_Long_CB = StringLen($Z_Zone)
$Z_Zone = ""

for $i = 0 to $Z_Long_CB-1
	if 	$aArray[$i] >= 48 and $aArray[$i] <= 57  or $aArray[$i] >= 65 and $aArray[$i] <= 90  or $aArray[$i] >= 97 and $aArray[$i] <= 122 or  $aArray[$i] = 32 Then
		$Z_Zone = $Z_Zone & chr($aArray[$i])
	Else
		$Z_Zone = $Z_Zone & " "
	EndIf
Next

Return $Z_Zone

EndFunc ; fin sup_car_sp


;----------------------------------------------------------------------------------------------------------------------
; envoie du mail
;----------------------------------------------------------------------------------------------------------------------
Func Envoie_Mail($Dest,$Cc,$Sujet,$Fichier_Attache,$Val_nom_table,$INSTANCE_SQL,$NOM_BDD)
    Local $olMailItem    = 0
    Local $olFormatRichText = 3
    Local $olImportanceLow   = 0
    Local $olImportanceNormal= 1
    Local $olImportanceHigh  = 2
    Local $olFormatHTML = 1

	Local $aRow, $SQL_err, $Z_Act, $sQuery
	Local  $aResult[10][4], $iRows,  $iColumns

_SQL_Startup()
_SQL_Connect(-1, $INSTANCE_SQL, $NOM_BDD, '', '')
if @error Then
	MsgBox(0, "ERROR", "Failed to connect to the database",10)
	Exit
EndIf

$sQuery = "SELECT COUNT(*)  FROM "& $Val_nom_table

consolewrite($sQuery)

$SQL_err = _SQL_GetTable2D(-1, $sQuery,  $aResult,  $iRows,  $iColumns)
if $SQL_err = 0 Then

	$sQuery = "set dateformat dmy ; "
	$sQuery = $sQuery & " INSERT INTO ["&$NOM_BDD&"].[dbo].[suivi_exploit]  ([Code_Table] ,[date_traitement] ,[nbr_lignes])  VALUES ( '" & $Val_nom_table & "', Getdate()," & $aResult[1][0] & " ) "

; mise a jour de la table de suivi d'exploit
	$SQL_err = _SQL_Execute(-1, $sQuery)

	if $SQL_err = $SQL_ERROR Then
		MsgBox (0, "erreur", "erreur insert suivi_exploit",10)
		Exit
	EndIf

	_SQL_Close()

EndIf
EndFunc