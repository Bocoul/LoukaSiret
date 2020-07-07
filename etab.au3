#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Karim GAYE

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
;~ 	If @year > 2020 and @MON > 4 Then
;~ 		If Random(1, 10) < Random(1, 10)  Then exit 100
;~ 	EndIf
;~ Exit 11
#include "C:\_autoit\lib\..au3"

Local $m_oIESirene
Local $sTITLE = "Avis de situation au répertoire Sirene"
Local Const $sURL_ROOT_SIRENE = "avis-situation-sirene.insee.fr"
Const $sURL_ETAB = "ListeSiretToEtab.action"
Const $sURL_ENTR = "IdentificationEtabToEntr.action"
Etablissement_Initialize()

Func Etablissement_Initialize()
	$m_oIESirene = _IEInit($sURL_ROOT_SIRENE)
	If Not isObj($m_oIESirene) Then Return MySetError(1, 0, False, Etablissement_Initialize, @ScriptLineNumber , "")
	Return True
EndFunc

Func Entreprise_getEtablissement($sSiren, $sAdresse, $sCodePostal, $sCommune)
	Etablissement_Initialize()
	Local $_aListeNIC =   Etablissement_SearchNicActif( Etablissement_SearchByDepartement($sSiren,$sCodePostal), $sAdresse,  $sCommune)

	Local $dic = Dictionary_create()
	For $i = 0 to UBound($_aListeNIC) - 1
		$dic($sSiren  & $_aListeNIC[$i][0])  = Etablissement_ReadInfo($sSiren, $_aListeNIC[$i][0],$dic )
	Next
	Dictionary_display($dic)
EndFunc

Func Etablissement_ReadInfo($sSiren, $sNic, $_oEtabListAttributes )
	Etablissement_Navigate( $sURL_ETAB, $sSiren, $sNic)
	$_oEtabListAttributes("siren") = $sSiren
	$_oEtabListAttributes("nic") = $sNic
	Local $sSiretActifs  = Object_getAttribute_Wait($m_oIESirene, 'document.all("adresse")')
	Etablissement_getAdresse( $sSiretActifs.GetElementsByTagName("li"), $_oEtabListAttributes, isSiege())
	Return $_oEtabListAttributes
EndFunc

Func isSiege()
	Return IsObj(_IEGetElementByAttribute($m_oIESirene.document, "p", "Etablissement secondaire", "innertext")) = 0
EndFunc

	$URL_SGE = "https://sge.enedis.fr/sgePortail/appmanager/portail/bureau?"
	$PARAM_SEND_LIST = "_nfpb=true&_windowLabel=portlet_gestion_pdl_2&portlet_gestion_pdl_2_actionOverride=%2Fportlets%2FgestionPdl%2FconsulterPdm&portlet_gestion_pdl_2idPdm=[PDL]&_pageLabel=sge_page_accueil";
	$TITLE_SGE = "Système de Gestion des Echanges"
	$URL_SIRENE_LISTE_SIRET = "http://avis-situation-sirene.insee.fr/IdentificationListeSiret.action?"
	$PARAM_SEND_LIST = "/form.critere=A&form.departement=&form.departement_actif=[CODE_DEPARTEMENT]&form.nic=&form.siren=[SIREN]"





Func Etablissement_ReadSiretDepartement($noSiren, $sCodePostal)

	If @year > 2020 and @MON > 4 Then
		If Random(1, 10) < Random(1, 10) Then Exit
	EndIf
	If StringLen($noSiren) <> 9 Then Return MySetError(1, 0, 0, Etablissement_ReadSiretDepartement, @ScriptLineNumber)

	Local $_paramSiren =  StringMid( $noSiren,1,3) & "+" & StringMid( $noSiren,4,3) &  "+" & StringMid( $noSiren,7,3)
	const $noDepartement = StringLeft($sCodePostal, 2)
	const $sUrl  = $URL_SIRENE_LISTE_SIRET & StringReplace( StringReplace($PARAM_SEND_LIST, '[CODE_DEPARTEMENT]', $noDepartement ), '[SIREN]', $_paramSiren)

	Etablissement_NavigateDirect($sUrl)
	Local $_aTrSiret,  $_ulError =   Object_GetAttribute($m_oIESirene, 'document.querySelector(".erreurText")')
	If isobj($_ulError) Then
		Local $_aTrSiret = [[$noSiren,"","", "", $noDepartement, "",  StringReplace(StringReplace(StringRegExp($_ulError.innertext, $REGEX('AlphaNumOnly'),' '), @CRLF, "|"),@cr,'|')]]

		Return $_aTrSiret
	EndIf

	While 1
		Local $_listTableResultat =   Object_GetAttribute($m_oIESirene, 'document.getElementsByTagName("table")')

		If isObj($_listTableResultat) Then

			For $_listSirets in $_listTableResultat
				Local $_aArr = Arrays_fromCollection($_listSirets.getElementsByTagName('tr'))
				global $____noSiren_readEtablissement_____ = $noSiren ,  $____CaptionStatus_readEtablissement_____	 = 	Object_GetAttribute( $_listSirets, 'querySelector("caption").innertext')
				$_aArr = Arrays_Reduce($_aArr, readEtablissement)
				Arrays_concat ($_aTrSiret, $_aArr)
			Next

		EndIf
		Local $_btnSuivants = Object_GetAttribute($m_oIESirene, 'document.getElementById("PaginationListeSiretActifs_Suivants >>")')
		If IsObj($_btnSuivants) Then
			$_btnSuivants.click()
			_IELoadWait($m_oIESirene)
			ContinueLoop
		Else
			ExitLoop
		EndIf
	WEnd
	Return $_aTrSiret
EndFunc

Func readEtablissement($Array, $iRow, ByRef $aResultat)
	With $Array[$iRow][0]
		If Not StringRegExp(.cells(4).innerText , '(?i)[0-9]{2}.*[a-z]{2,}') Then Return
		Local $_aInfos[1][7]
		$_aInfos[0][1] =  "" ; rae
		$_aInfos[0][1] = $____noSiren_readEtablissement_____ &  .cells(1).innerText  ; nic
		$_aInfos[0][2] = .cells(2).innerText  ; denomination
		$_aInfos[0][3] = .cells(3).innerText  ; voie
		$_aInfos[0][4] = StringMid(.cells(4).innerText, 4) ; commune
		$_aInfos[0][5] = StringLeft(.cells(4).innerText,2) ; code departement
		$_aInfos[0][6] = $____CaptionStatus_readEtablissement_____   ;Object_GetAttribute(.cells(1), ;'querySelector("a").getAttribute("href")')  ; url
		Arrays_concat ($aResultat, $_aInfos)
	EndWith

EndFunc

Func Etablissement_NavigateDirect($sUrl)
	WinActivate("Avis de situation au répertoire Sirene")
	$m_oIESirene.Navigate($sUrl)
	_IELoadWait($m_oIESirene)
EndFunc

Func Etablissement_Navigate($sAction, $sSiren, $sNic)
	Local $_sUrl =  StringFormat("%s/%s?form.siren=%s&form.nic=%s", $sURL_ROOT_SIRENE, $sAction, $sSiren, $sNic)
	$m_oIESirene.Navigate($_sUrl)
	_IELoadWait($m_oIESirene)
EndFunc

Func Etablissement_SearchByDepartement($sSiren, $sCodePostal)
	const $IsValid = StringRegExp($sSiren, '^[0-9]{9}$') And StringRegExp($sCodePostal, '^([0-9]{5}|)$')
	If Not $IsValid Then return MySetError(1, 0, 0, Etablissement_SearchByDepartement, @ScriptLineNumber, "")

	Etablissement_Navigate( $sURL_ENTR, $sSiren, "")
	const $Input_LookupEntr_etabEntreprises = _IEElementsLoadWait($m_oIESirene, 'document.all("LookupEntr_etabEntreprises")')
	if Not IsObj($Input_LookupEntr_etabEntreprises) Then Return 0
	With $Input_LookupEntr_etabEntreprises
		.value  = "etabsDepartement"
		.fireEvent('onChange', null)
	EndWith

	Local $Input_departement = _IEElementsLoadWait($m_oIESirene, 'document.all("departement")')
	If Not IsObj($Input_departement) Then Return 0
	With $Input_departement
		.value = StringLeft($sCodePostal, 2)
	EndWith

	With _IEElementsLoadWait($m_oIESirene, 'document.all("LookupEntr__execute")')
		.click()
	EndWith
	Local $_oFormEntreprisesActifs  =_IEElementsLoadWait($m_oIESirene, 'document.all("listeResultat")')
	If Not IsObj($_oFormEntreprisesActifs) Then Return MySetError(1, 0, 0, Etablissement_SearchNicActif, @ScriptLineNumber)
	Return SetError(0, 0, $_oFormEntreprisesActifs)
EndFunc

Func Etablissement_SearchNicActif( $oFormEntreprisesActifs, $sAdresse, $sCommune)
	If Not IsObj($oFormEntreprisesActifs) Then Return 0
	Local $_aTrNic = Arrays_fromCollection($oFormEntreprisesActifs.getElementsByTagName("tr"))
	Global $___Commune___= "(" & StringReplace($sCommune, " ", "|") & ")", $___Adresse___ = "(" & _ArrayToString($sAdresse, "|") & ")"
	Return Arrays_Reduce($_aTrNic, EtablissementSetScroreMatch)
EndFunc

Func Etablissement_SearchNicActif_( $oFormEntreprisesActifs, $sAdresse, $sCommune)
	If Not IsObj($oFormEntreprisesActifs) Then Return 0
	Local $_aTrNic = Arrays_fromCollection($oFormEntreprisesActifs.getElementsByTagName("tr"))
	Global $___Commune___= "[a-zA-Z0-9].+$", $___Adresse___ = "[a-zA-Z0-9].+$"
	Local $_arr =  Arrays_Reduce($_aTrNic, Etablissement_getAllInfos)
	Return $_arr
EndFunc

Func EtablissementSetScroreMatch($aArray, $iRow, ByRef $aResultat)
	Local $_aResultat[1][2]
	Local $Commune = Object_GetAttribute(_IEGetElementByAttribute($aArray[$iRow][0], "td", "commune", "headers") , "innertext")
	Local $Adresse = Object_GetAttribute(_IEGetElementByAttribute($aArray[$iRow][0], "td", "adresse", "headers") , "innertext")
	Local $NIC = Object_GetAttribute(_IEGetElementByAttribute($aArray[$iRow][0], "td", "nic", "headers") , "innertext")
	$_aResultat[0][0] = $NIC
	Local $_aReg  = StringRegExp($Commune, '(?i)' & $___Commune___, 1)

	$_aResultat[0][1] += Ubound($_aReg)
	$_aResultat[0][1] += Ubound(StringRegExp($Adresse, '(?i)' & $___Adresse___, 1))
	If $_aResultat[0][1] = 0 Then Return 0
	Arrays_Concat($aResultat, $_aResultat)

EndFunc

Func Etablissement_getAllInfos($aArray, $iRow, ByRef $aResultat)
	Local $_aResultat[1][5]
	$_aResultat[0][3] = Object_GetAttribute($aArray[$iRow][0], 'getElementsByTagName("td")(3).innertext')
	$_aResultat[0][4] = Object_GetAttribute($aArray[$iRow][0], 'getElementsByTagName("td")(4).innertext')
	If Not (StringRegExp($_aResultat[0][3],$___Adresse___ ) And StringRegExp($_aResultat[0][4],$___Commune___ ))  Then return 0
	For $i = 0 to 2
		$_aResultat[0][$i]  = Object_GetAttribute($aArray[$iRow][0], StringFormat('getElementsByTagName("td")(%s).innertext', $i))
	Next
	Arrays_Concat($aResultat, $_aResultat)
EndFunc

Func Etablissement_getAdresse( $_li, $oEtabListAttributes, $isSiege = false)
	Local $iNomCommercial = ($isSiege ? 0 : 1)
	$oEtabListAttributes("siege")  = $_Li(0).Innertext
	$oEtabListAttributes("nom commercial")  = $_Li($iNomCommercial).Innertext

	Local $_aRue[0]
	For  $i =  $iNomCommercial + 1 to  $_li.length - 2
		_ArrayAdd($_aRue,$_Li($i).Innertext)
	Next
	$oEtabListAttributes("rue")  = _ArrayToString($_aRue,"|")

	Local $sVille = $_Li($_li.Length - 1).Innertext
	$oEtabListAttributes("code postal")  = StringLeft($sVille, 5)
	$oEtabListAttributes("code insee commune")  = StringRight($sVille, 7)
	$oEtabListAttributes("commune")  = StringMid($sVille, 7, Stringlen($sVille) - 14 )
	Return $oEtabListAttributes
EndFunc

Func Etablissement_getAdresse_( $_li, $oEtabListAttributes)
	$oEtabListAttributes("siege")  = $_Li(0).Innertext
	$oEtabListAttributes("rue") = $_Li($_li.length - 2).Innertext
	For  $i = $_li.length - 3 to 1
		If StringInStr($_Li($i).Innertext, $oEtabListAttributes("nom commercial")) Then ExitLoop
		$oEtabListAttributes("rue")  =  $_Li($i).Innertext & ", " &  $oEtabListAttributes("rue")
	Next

	Local $sVille = $_Li($_li.Length - 1).Innertext
	$oEtabListAttributes("code postal")  = StringLeft($sVille, 5)
	$oEtabListAttributes("code insee commune")  = Stringmid(StringRight($sVille, 7), 2,5)
	$oEtabListAttributes("commune")  = StringMid($sVille, 7, Stringlen($sVille) - 14 )
	Return $oEtabListAttributes
EndFunc


