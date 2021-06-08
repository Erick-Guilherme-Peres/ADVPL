#include 'protheus.ch'
#include "topconn.ch"

/************************************************************************************************/
/* Rotina: TFFINP08																				*/
/* Desenvolvimento: Erick Guilherme Peres		Validado Por: Erick Guilherme Peres				*/	
/* Data Cria��o: 21/12/200						Data Valida��o:	19/02/2021						*/
/* Funcionalidade: Rotina automatica para retornar a cota��o das moedas do Banco Central		*/
/*																								*/
/*|Vers�o | Observa��es															   | Valida��o |*/
/*|-------|------------------------------------------------------------------------|-----------|*/
/*| 001.0 |	Cria��o da Funcionalidade											   | 21/12/2020|*/
/************************************************************************************************/

User Function TFFINP08( aParam ) // Recebo como Parametros os dados padr�es do Cadastro da Tarefa no WorkFlow ( Empresa, Filial, Unidade, etc.. )
Local aMoedas := {}
Local aRetMoeda

Local cHtml
Local cEmpresa
Local cFilEmp
Local cBase

Local cDest := ""	// Define o(s) Destinat�rio(s) do E-mail
Local dDataRef

Local nCont
Local nMoeda2
Local nTentativa

Local lTable
Local lMoeda := .T.
Local nX := 1

	if ( Select("SX2") == 0 )
		cEmpresa := aParam[1]
		cFilEmp  := aParam[2]
		
		RPCSetType(3)
		RPCSetEnv( cEmpresa, cFilEmp,"","","","",{"SM0", "SM2"})
	endIf

	aMoedas := MoedaSist()	//Chama Fun��o para retornar as moedas que a empresa utiliza
	
	cBase := Substr(AllTrim( SuperGetMv( "MV_X_APLIC",, "" ) ), 1,1)   //Par�metro Customizado que define ambiente de execu��o da Aplica��o do Protheus
	cHtml := '<html>'
	cHtml += '	<head>'
	cHtml += '		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />'
	cHtml += '		<title>Cota��o Monet�ria Dia: ' + DToC( dDataBase ) + '</title>'
	cHtml += '		<style type="text/css">' 
	cHtml += '			<!-- body {background-color: #9AB986;} 2Estyle1 {font-family: Segoe UI,Verdana, Arial;font-size: 12pt;}'
	cHtml += '								2Estyle2 {font-family: Segoe UI,Verdana, Arial;font-size: 12pt;color: rgb(255,0,0)}' 
	cHtml += '								2Estyle3 {font-family: Segoe UI,Verdana, Arial;font-size: 10pt;color: rgb(37,64,97)}' 
	cHtml += '								2Estyle4 {font-size: 8pt; color: rgb(37,64,97); font-family: Segoe UI,Verdana, Arial;}'
	cHtml += '								2Estyle5 {font-size: 10pt} --> '  
	cHtml += '		</style>'
	cHtml += '	</head>'
	cHtml += '		<body style="background-color: #9AB986;">'
	cHtml += '			<table align="center"; style="background-color: rgb(255, 255, 255); width: 520px; text-align: left; margin-left: auto; margin-right: auto;" id="total" border="0" cellpadding="12">' 
	cHtml += '			<tbody>'
		
	if cBase == "H"
		cHtml += '					<tr> '     
		cHtml += '						<td colspan="4"> '     
		cHtml += '							<p class="style1">'
		cHtml += '							Ambiente de homologa��o</p>'
		cHtml += '						</td>'    
		cHtml += '					</tr>'
		
		//Quando for Homologa��o, por padr�o altera o(s) destinat�rio(s) do e-mail, para n�o enviar por exemplo ao Depto Financeiro sem a devida necessidade
		cDest := "email@destinatario.com.br"	//E-mail do Destinat�rio
	endif
			
	cHtml += '					<tr> '     
	cHtml += '						<td colspan="4"> '     
	cHtml += '							<p class="style1">'
	cHtml += '								Esta mensagem refere-se a Cota��o Monet�ria da ' + AllTrim( Capital( FwGrpName() ) ) + ': '
	cHtml += '							</p>'
	cHtml += '						</td>'    
	cHtml += '					</tr>'       					 
	cHtml += '					<tr> '     
	cHtml += '						<td colspan="4">'
	cHtml += "							<table align='center' style='background-color: rgb(240, 240, 240);width: 100%; text-align: left;' id='total' border='0' cellpadding='12'>"
	cHtml += "								<tr>"
	cHtml += "									<td colspan='4' bgcolor='#F7F7F7' style='font-size: 13px;'>"
	cHtml += "										<strong>Data Sistema</strong>"
	cHtml += "									</td>"
	cHtml += "									<td colspan='2' bgcolor='#F7F7F7' style='font-size: 13px;'>"
	cHtml += "										<strong>Moeda</strong>"
	cHtml += "									</td>"
	cHtml += "									<td colspan='3' bgcolor='#F7F7F7' style='font-size: 13px;'>"
	cHtml += "										<strong>Data Banco Central</strong>"
	cHtml += "									</td>"
	cHtml += "									<td colspan='3' bgcolor='#F7F7F7' style='font-size: 13px;'>"
	cHtml += "										<strong>Valor</strong>"
	cHtml += "									</td>"		
	cHtml += "								</tr>"
	
	for nCont := 1 to Len( aMoedas ) //Percorre Array com as Moedas que a empresa Trabalha
		Do Case
		 Case Dow( dDataBase ) == 1
		 	dDataRef := DaySub( dDataBase, 2)
		 Case Dow( dDataBase ) == 2
		 	dDataRef := DaySub( dDataBase, 3)	 
		 OTHERWISE 
		 	dDataRef := DaySub( dDataBase, 1)
		EndCase
		
		nMoeda2 := 0
		nTentativa := 0
		lMoeda := .T.
		while lMoeda
			aRetMoeda := CapturaMoeda( dDataRef, aMoedas[ nCont, 2 ] ) //Chama Fun��o que realizar� a Captura do Valor da Cota��o
			lMoeda	:= !aRetMoeda[1]
			nMoeda2 := aRetMoeda[2]
			nTentativa++
			
			if nTentativa > 10
				lMoeda := .F.
				nMoeda2 := 0
			endif
		end
		
		if !(nTentativa == 11 .And. nMoeda2 == 0) .Or. nMoeda2 > 0
			dbSelectArea("SM2")
			dbSetOrder(1)
			lTable := !dbSeek( DToS( dDataBase ) )
			
			if RecLock("SM2", lTable )
				SM2->M2_DATA 	:= dDataBase				
				SM2->M2_INFORM	:= "S"
				&("SM2->M2_MOEDA" + cValToChar( aMoedas[ nCont, 3 ] ) )	:= nMoeda2
				
				MsUnlock()
			endif
			
			cHtml += "								<tr>"
			cHtml += "									<td colspan='4' bgcolor='#F7F7F7' style='font-size: 11px;'>" + DToc( dDataBase ) + "</td>"
			cHtml += "									<td colspan='2' bgcolor='#F7F7F7' style='font-size: 11px;'>" + aMoedas[ nCont, 4 ] + "</td>"
			cHtml += "									<td colspan='3' bgcolor='#F7F7F7' style='font-size: 11px;'>" + DToC( dDataRef ) + "</td>"
			cHtml += "									<td colspan='3' bgcolor='#F7F7F7' style='font-size: 11px;'>" + Transform( nMoeda2, "@e 999,999.9999" ) + "</td>"
			cHtml += "								</tr>"
			
		    if Dow( dDataBase ) == 7
			    For nX:=1 to 2	
			    	dDataBase:= DaySum(dDataBase,1)	
			    	lTable := !dbSeek( DToS( dDataBase ) )
					if RecLock("SM2", lTable )
						SM2->M2_DATA 	:= dDataBase		
						SM2->M2_INFORM	:= "S"
						&("SM2->M2_MOEDA" + cValToChar( aMoedas[ nCont, 3 ] ) )	:= nMoeda2
						
						MsUnlock()
					endif									    	
			    	
					cHtml += "							<tr>"
					cHtml += "								<td colspan='4' bgcolor='#F7F7F7' style='font-size: 11px;'>" + DToc( dDataBase ) + "</td>"
					cHtml += "								<td colspan='2' bgcolor='#F7F7F7' style='font-size: 11px;'>" + aMoedas[ nCont, 4 ] + "</td>"
					cHtml += "								<td colspan='3' bgcolor='#F7F7F7' style='font-size: 11px;'>" + DToC( dDataRef ) + "</td>"
					cHtml += "								<td colspan='3' bgcolor='#F7F7F7' style='font-size: 11px;'>" + Transform( nMoeda2, "@e 999,999.9999" ) + "</td>"
					cHtml += "							</tr>"	   
			    Next	
			    dDataBase := DaySub( dDataBase, 2)  	
		    endif			
		endif
    next

    cHtml += "							</table>"
	cHtml += "						</td>"
	cHtml += "					</tr>"
	cHtml += "					<tr>"
	cHtml += "						<td colspan='4' class='style4'>"
	cHtml += "							<span class='style5'>"
	cHtml += "								<em>"
	cHtml += "									<span style='font-size: 10px;text-decoration: underline;'> powered by Toffano Produtos Alimenticios Ltda. - Departamento de TI</span>"
	cHtml += "								</em>"
	cHtml += "								<em></em>"
	cHtml += "							</span>"
	cHtml += "						</td>"
	cHtml += '					</tr> '  
	cHtml += '			</tbody>'
	cHtml += '		</table>'
	cHtml += '	</body>'
	cHtml += '</html>'
	
	U_EnviarEMail( cDest, "[FINANCEIRO] Cota��o Monet�ria ( " + AllTrim( Capital( FwGrpName() ) ) + " ) - Dia: "  + DToC( dDataBase ), cHtml, "", .f., .F. )  // Fun��o customizada para disparar o e-mail confirmando as altera��es da Cota��o da Moeda 
Return

Static Function MoedaSist() // Fun��o que Retorna Array Com Simbolo, Codigo do Banco Central, Posicao da Tabela SM2 e Descricao das Moedas Estrangeiras 
Local aMoedas 		:= {}
Local aMoedasSist	:= {}

Local nCont
Local nMoedas
Local cTitulo

	aAdd( aMoedas, { "US$", 1 } ) 		// Array com o Simbolo da Moeda e o C�digo da Moeda no Banco Central
	aAdd( aMoedas, { "EURO", 21619 } )	// o C�digo da Moeda no Banco Central, voc� descobre atrav�s do seguinte endere�o: https://egas.digital/cotacoes.txt

	for nCont := 2 to 99
		if GetMv( "MV_SIMB" + cValToChar( nCont ), .T. )
			nMoedas := Ascan( aMoedas, { |x| AllTrim( x[1] ) == AllTrim( GetMv( "MV_SIMB" + cValToChar( nCont ) ) ) } )
			cTitulo := AllTrim( GetMv( "MV_MOEDA" + cValToChar( nCont ) ) )
			
			if nMoedas != 0
				nMoedas := aMoedas[ nMoedas, 2 ]
				aAdd( aMoedasSist, { AllTrim( GetMv( "MV_SIMB" + cValToChar( nCont ) ) ), nMoedas, nCont, cTitulo } )
			endif
		endif
	next
Return aMoedasSist

Static Function CapturaMoeda( dDataRef, nMoeda )	//Fun��o que realiza a integra��o com a API do Banco Central e Retorna o Valor da Cota��o, conforme Data Informada
Local aHeader := {}

Local cURL  := "https://api.bcb.gov.br"
Local cPath

Local oRest
Local oJson

Local nValor := 0

Local lRet := .F.

	aadd(aHeader,'User-Agent: Mozilla/4.0 (compatible; Protheus '+GetBuild()+')')
    aAdd(aHeader,'Content-Type: application/json; charset=utf-8')
	
	cPath	:= "/dados/serie/bcdata.sgs." + cValToChar( nMoeda ) + "/dados?dataInicial=" + DToC( dDataRef ) + "&dataFinal=" + DToC( dDataRef ) + ""
	
	oRest := FwRest():New(cURL)
	oRest:setPath( cPath )

	If oRest:Get(aHeader)		
		oJson := JsonObject():new() 
		oJson:FromJson( StrTran(STrTran(oRest:cResult, '[', ''),']', '' ) )
		nValor := Val( oJson:GetJsonText("valor") )
		
		if nValor != 0	//Se Valor For Diferente de 0, sistema Obteve exito em capturar a Moeda. Essa tratativa foi adicionada pois podem ocorrer erros na captura do valor.
			lRet := .T.
		endif
	endif
Return { lRet, nValor }