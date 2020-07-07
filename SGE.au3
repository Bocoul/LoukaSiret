#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Karim GAYE

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

#include "C:\_autoit\lib\..au3"

Local $m_oIESGE
Local $sTITLE = "Avis de situation au répertoire Sirene"
Local Const $sURL_SGE_ROOT = "https://sge.enedis.fr"
Const $sURL_PDL = "/sgePortail/appmanager/portail/bureau?_nfpb=true&_windowLabel=portlet_gestion_pdl_2&portlet_gestion_pdl_2_actionOverride=%2Fportlets%2FgestionPdl%2FconsulterPdm&portlet_gestion_pdl_2idPdm=[PDL]&_pageLabel=sge_page_accueil"


SGE_Initialize()
Func SGE_ReadAdressePDL($IdPDL)
	If StringLen($IdPDL) <> 14 Then Return MySetError(1, 0, 0, SGE_ReadAdressePDL, @ScriptLineNumber)
	While Not SGE_Initialize()
		Sleep(100)
	WEnd
	SGE_Navigate($IdPDL)
	Local $_oDocHtml = _IEElementsLoadWait($m_oIESGE, "document.body")
	Local $_iParagraphLength = _IEElementsLoadWait($_oDocHtml, 'getElementsByTagName("p").length')
	Local $_oParagraph_Gen  = _IEGetElementByAttributewait( $_oDocHtml, "p",$sTitle, "innertext")

	$_oPrmId =  object_getAttribute( $_oDocHtml, 'querySelector("#prmId")')
	If IsObj($_oPrmId) Then
		With $_oDocHtml
			Local $_aInfos =  [["", _
								$_oPrmId.value, _  ; pdl
								.all("voie").value, _ ; voie
								.all("codePostal").value, _ ; code postal
								.all("commune").value, _ ;commune
								StringLeft(.all("codePostal").value, 2), _ ; departement
								.all("codeINSEE").value]] ; code insee
			Return $_aInfos
		EndWith
	EndIf

EndFunc

Func SGE_getDivParagraph($IdPDL, $sTitle =  "Données générales" )
	While Not SGE_Initialize()
		Sleep(100)
	WEnd
	SGE_Navigate($IdPDL)
	Local $_oDocHtml = _IEElementsLoadWait($m_oIESGE, "document.body")
	Local $_iParagraphLength = _IEElementsLoadWait($_oDocHtml, 'getElementsByTagName("p").length')
	Local $_oParagraph_Gen  = _IEGetElementByAttributewait( $_oDocHtml, "p",$sTitle, "innertext")


	Local $oBlocAddress = _IEElementsLoadWait($_oParagraph_Gen, "nextSibling")
	Return $oBlocAddress
EndFunc

Func SGE_Initialize()
	$m_oIESGE = _IEInit($sURL_SGE_ROOT)
	If Not isObj($m_oIESGE) Then Return MySetError(1, 0, False, SGE_Initialize, @ScriptLineNumber , "")
	Return True
EndFunc

Func SGE_Navigate($IdPDL)
WinActivate("Système de Gestion des Echanges")

	Local $_sUrl =  StringReplace($sURL_SGE_ROOT & $sURL_PDL, "[PDL]", $IdPDL)
	$m_oIESGE.Navigate($_sUrl)

	_IELoadWait($m_oIESGE)
EndFunc