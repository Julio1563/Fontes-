#include "protheus.ch"
#include "topconn.ch"
#include "rwmake.ch"

/*/{Protheus.doc} NSAR002

@type function
@version 12.1.2210
@author Danilo Azevedo
@since 03/12/2019
@return variant, return_description
/*/
User Function Testej2()

	Local oReport
	Local aRet 		:= {}
	Local aParamBox := {}

	Private _cAlias1	:= ""
	Private dDtIni
	Private dDtFim
	Private dDigIni
	Private dDigFim

	aadd(aParamBox, {1, "Emissão de"   , ctod(space(8)), "", ".T.", "", "", 50, .T.}) // Tipo caractere
	aadd(aParamBox, {1, "Emissão até"  , ctod(space(8)), "", ".T.", "", "", 50, .T.}) // Tipo caractere
	aadd(aParamBox, {1, "Digitação de" , ctod(space(8)), "", ".T.", "", "", 50, .T.}) // Tipo caractere
	aadd(aParamBox, {1, "Digitação até", ctod(space(8)), "", ".T.", "", "", 50, .T.}) // Tipo caractere
	If !ParamBox(aParamBox,"NSAR002 - Parâmetros...",@aRet)
		Return()
	Endif

	dDtIni  := aRet[1]
	dDtFim  := aRet[2]
	dDigIni := aRet[3]
	dDigFim := aRet[4]

	_cAlias1 := GetNextAlias()

	oReport := ReportDef()
	oReport:PrintDialog()

Return()

/*
Funcao: ReportDef()
Descricao: Cria a estrutura do relatorio
*/

Static Function ReportDef(cPerg)

	Local oReport
	Local oSection1
	Local aOrdem	:= {}

	//Declaracao do relatorio
	oReport := TReport():New("NSAR002","NSAR002 - Análise Contábil",,{|oReport| PrintReport(oReport)},"NSAR002 - Análise Contábil")

	oReport:PrintHeader(.T.,.T.)
	//Ajuste nas definicoes
	oReport:nLineHeight := 55
	oReport:cFontBody 	:= "Courier New"
	oReport:nFontBody 	:= 12		//&& 10
	oReport:lHeaderVisible := .T.
	oReport:lDisableOrientation := .T.
	//oReport:SetLandscape()
	oReport:SetPortrait()

	oReport:SetTotalInLine(.F.)
	oReport:SetTitle('Análise Contábil')
	//oReport:SetLineHeight(30)
	oReport:SetColSpace(1)
	//oReport:SetLeftMargin(0)
	//oReport:oPage:SetPageNumber(1)
	oReport:lBold := .T.
	oReport:lUnderLine := .F.
	cBitmap := alltrim("LGRL"+SM0->M0_CODIGO+SM0->M0_CODFIL)+".BMP" // Empresa+Filial
	_cFileLogo	:= GetSrvProfString('Startpath','') + cBitmap
	//oReport:SayBitmap(25,25,_cFileLogo,150,60) // insere o logo no relatorio

	//Secao do relatorio
	oSection1 := TRSection():New(oReport,"Lançamentos Contábeis",{_cAlias1,"NAOUSADO"},aOrdem,,,,,,.T.,.T.,.T.)

	//Celulas da secao D1_FILIAL, D1_EMISSAO, D1_DTDIGIT, D1_DOC, A2_NOME, D1_TOTAL, D1_VALIRR, D1_VALPIS, D1_VALCOF, D1_VALCSL, D1_VALINS, D1_VALISS,
	//E2_VALLIQ, E2_VALOR, E2_CCC, E2_CCD, E2_CCUSTO, isnull(DE_CONTA,'') DE_CONTA, isnull(DE_ITEMCTA,'') DE_ITEMCTA, isnull(DE_CUSTO1,0) DE_CUSTO1
	TRCell():New(oSection1, "D1_FILIAL" , "SD1")
	TRCell():New(oSection1, "D1_EMISSAO", "SD1")
	TRCell():New(oSection1, "D1_DTDIGIT", "SD1")
	TRCell():New(oSection1, "D1_DOC"    , "SD1")
	TRCell():New(oSection1, "A2_NOME"   , "SA2")
	TRCell():New(oSection1, "D1_ITEM"   , "SD1")
	TRCell():New(oSection1, "D1_COD"    , "SD1")
	TRCell():New(oSection1, "D1_TOTAL"  , "SD1")
	TRCell():New(oSection1, "D1_VALIRR" , "SD1")
	TRCell():New(oSection1, "D1_VALPIS" , "SD1")
	TRCell():New(oSection1, "D1_VALCOF" , "SD1")
	TRCell():New(oSection1, "D1_VALCSL" , "SD1")
	TRCell():New(oSection1, "D1_VALINS" , "SD1")
	TRCell():New(oSection1, "D1_VALISS" , "SD1")

	TRCell():New(oSection1, "D1_BASIMP5", "SD1")
	TRCell():New(oSection1, "D1_ALQIMP5", "SD1")
	TRCell():New(oSection1, "D1_VALIMP5", "SD1")
	TRCell():New(oSection1, "D1_BASIMP6", "SD1")
	TRCell():New(oSection1, "D1_ALQIMP6", "SD1")
	TRCell():New(oSection1, "D1_VALIMP6", "SD1")

	TRCell():New(oSection1, "E2_VALLIQ" , "SE2")
	TRCell():New(oSection1, "E2_VALOR"  , "SE2")
	TRCell():New(oSection1, "E2_CCC"    , "SE2")
	TRCell():New(oSection1, "E2_CCD"    , "SE2")
	TRCell():New(oSection1, "DE_CC"     , "SDE")
	TRCell():New(oSection1, "DE_CONTA"  , "SDE")
	TRCell():New(oSection1, "DE_ITEMCTA", "SDE")
	TRCell():New(oSection1, "DE_CUSTO1" , "SDE")
	oSection1:SetPageBreak(.F.)

Return oReport

Static Function PrintReport(oReport)

	Local oSection1 := oReport:Section(1)
	Local _cQuery   := ""

	oReport:oPage:setPaperSize(9)
	oBrush := TBrush():New( , CLR_BLACK )

	/*
	If oreport:OREPORT:NDEVICE <> 4
	MsgInfo("Atenção: este relatório foi formatado para Excel. Os demais formatos podem não apresentar o layout correto.")
	Endif
	*/

	/*
	select DE_CONTA, DE_ITEMCTA, DE_CUSTO1 from SD1010 D1
	left join SDE010 DE on D1_FILIAL = DE_FILIAL and D1_DOC = DE_DOC and D1_SERIE = DE_SERIE and D1_FORNECE = DE_FORNECE and D1_LOJA = DE_LOJA and DE.D_E_L_E_T_ = ''
*/

	_cQuery := "select D1_FILIAL, D1_EMISSAO, D1_DTDIGIT, D1_DOC, A2_NOME, D1_ITEM, D1_COD, D1_TOTAL, D1_VALIRR, D1_VALPIS, D1_VALCOF, D1_VALCSL, D1_VALINS, D1_VALISS , E2_VALLIQ, D1_BASIMP5, D1_ALQIMP5, D1_VALIMP5, D1_BASIMP6, D1_ALQIMP6, D1_VALIMP6, E2_VALOR, E2_CCC, E2_CCD, E2_CCUSTO, isnull(DE_CC,'') DE_CC, isnull(DE_CONTA,'') DE_CONTA, isnull(DE_ITEMCTA,'') DE_ITEMCTA, isnull(DE_CUSTO1,0) DE_CUSTO1
	_cQuery += " from "+RetSqlName("SD1")+" SD1
	_cQuery += " left join "+RetSqlName("SDE")+" SDE on D1_FILIAL = DE_FILIAL and D1_DOC = DE_DOC and D1_SERIE = DE_SERIE and D1_FORNECE = DE_FORNECE and D1_LOJA = DE_LOJA and D1_ITEM = DE_ITEMNF and SDE.D_E_L_E_T_ = ''
	_cQuery += " join "+RetSqlName("SA2")+" SA2 on D1_FORNECE = A2_COD and D1_LOJA = A2_LOJA
	_cQuery += " join "+RetSqlName("SE2")+" SE2 on D1_FILIAL = E2_FILORIG and D1_DOC = E2_NUM and D1_SERIE = E2_PREFIXO and D1_FORNECE = E2_FORNECE and D1_LOJA = E2_LOJA
	_cQuery += " where A2_FILIAL = '"+xFilial("SA2")+"'"
	_cQuery += " and D1_FILIAL = '"+xFilial("SF1")+"'"
	_cQuery += " and E2_FILIAL = '"+xFilial("SE2")+"'"
	_cQuery += " and D1_EMISSAO between '"+dtos(dDtIni)+"' and '"+dtos(dDtFim)+"'
	_cQuery += " and D1_DTDIGIT between '"+dtos(dDigIni)+"' and '"+dtos(dDigFim)+"'
	_cQuery += " and D1_TIPO = 'N'
	_cQuery += " and SD1.D_E_L_E_T_ = ''
	_cQuery += " and SE2.D_E_L_E_T_ = ''
	_cQuery += " and SA2.D_E_L_E_T_ = ''
	_cQuery += " order by D1_FILIAL, D1_EMISSAO, D1_FORNECE, D1_LOJA, D1_DOC, D1_SERIE, D1_ITEM
	DbUseArea(.T., 'TOPCONN', TCGenQry(,,_cQuery), _cAlias1, .F., .T.)
	TCSetFiEld(_cAlias1,"D1_EMISSAO","D",8,0)
	TCSetFiEld(_cAlias1,"D1_DTDIGIT","D",8,0)

	Do While !(_cAlias1)->(Eof())
		If oReport:Cancel()
			Exit
		EndIF

		oReport:IncMeter()
		oSection1:Init()

		oSection1:Cell("D1_FILIAL"):SetValue((_cAlias1)->D1_FILIAL)
		oSection1:Cell("D1_EMISSAO"):SetValue((_cAlias1)->D1_EMISSAO)
		oSection1:Cell("D1_DTDIGIT"):SetValue((_cAlias1)->D1_DTDIGIT)
		oSection1:Cell("D1_DOC" ):SetValue((_cAlias1)->D1_DOC)
		oSection1:Cell("A2_NOME" ):SetValue((_cAlias1)->A2_NOME)
		oSection1:Cell("D1_ITEM"):SetValue((_cAlias1)->D1_ITEM)
		oSection1:Cell("D1_COD"):SetValue((_cAlias1)->D1_COD)
		oSection1:Cell("D1_TOTAL"):SetValue((_cAlias1)->D1_TOTAL)
		oSection1:Cell("D1_VALIRR" ):SetValue((_cAlias1)->D1_VALIRR)
		oSection1:Cell("D1_VALPIS" ):SetValue((_cAlias1)->D1_VALPIS)
		oSection1:Cell("D1_VALCOF"):SetValue((_cAlias1)->D1_VALCOF)
		oSection1:Cell("D1_VALCSL"):SetValue((_cAlias1)->D1_VALCSL)
		oSection1:Cell("D1_VALINS" ):SetValue((_cAlias1)->D1_VALINS)
		oSection1:Cell("D1_VALISS" ):SetValue((_cAlias1)->D1_VALISS )

		oSection1:Cell("D1_BASIMP5" ):SetValue((_cAlias1)->D1_BASIMP5)
		oSection1:Cell("D1_ALQIMP5" ):SetValue((_cAlias1)->D1_ALQIMP5)
		oSection1:Cell("D1_VALIMP5" ):SetValue((_cAlias1)->D1_VALIMP5)
		oSection1:Cell("D1_BASIMP6" ):SetValue((_cAlias1)->D1_BASIMP6)
		oSection1:Cell("D1_ALQIMP6" ):SetValue((_cAlias1)->D1_ALQIMP6)
		oSection1:Cell("D1_VALIMP6" ):SetValue((_cAlias1)->D1_VALIMP6)

		oSection1:Cell("E2_VALLIQ" ):SetValue((_cAlias1)->E2_VALLIQ)
		oSection1:Cell("E2_VALOR" ):SetValue((_cAlias1)->E2_VALOR)
		oSection1:Cell("E2_CCC" ):SetValue((_cAlias1)->E2_CCC)
		oSection1:Cell("E2_CCD" ):SetValue((_cAlias1)->E2_CCD)
		oSection1:Cell("DE_CC" ):SetValue((_cAlias1)->DE_CC)
		oSection1:Cell("DE_CONTA" ):SetValue((_cAlias1)->DE_CONTA)
		oSection1:Cell("DE_ITEMCTA" ):SetValue((_cAlias1)->DE_ITEMCTA)
		oSection1:Cell("DE_CUSTO1" ):SetValue((_cAlias1)->DE_CUSTO1)

		oSection1:PrintLine()
		oReport:SkipLine()
		(_cAlias1)->(dbSkip())

	EndDo
	oSection1:Finish()

	If Select(_cAlias1) > 0
		dbSelectArea(_cAlias1)
		dbCloseArea()
	EndIf

Return()
