/*
 * Harbour Project source code:
 *
 * Copyright 2011 Fausto Di Creddo Trautwein, ftwein@yahoo.com.br
 * www - http://harbour-project.org
 *
 * Thanks TO Robert F Greer, PHP original version
 * http://sourceforge.net/projects/excelwriterxml/
 *
 * This program is free software; you can redistribute it AND/OR modify
 * it under the terms of the GNU General PUBLIC License as published by
 * the Free Software Foundation; either version 2, OR( at your option )
 * any later version.
 *
 * This program is distributed IN the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General PUBLIC License FOR more details.
 *
 * You should have received a copy of the GNU General PUBLIC License
 * along WITH this software; see the file COPYING.txt.  IF NOT, write TO
 * the Free Software Foundation, Inc., 59 Temple Place, Suite 330,
 * Boston, MA 02111-1307 USA( OR visit the web site http://www.gnu.org/ ).
 *
 * As a special exception, the Harbour Project gives permission FOR
 * additional uses of the text contained IN its release of Harbour.
 *
 * The exception is that, IF you link the Harbour libraries WITH other
 * files TO produce an executable, this does NOT by itself cause the
 * resulting executable TO be covered by the GNU General PUBLIC License.
 * Your use of that executable is IN no way restricted on account of
 * linking the Harbour library code into it.
 *
 * This exception does NOT however invalidate any other reasons why
 * the executable file might be covered by the GNU General PUBLIC License.
 *
 * This exception applies only TO the code released by the Harbour
 * Project under the name Harbour.  IF you copy code FROM other
 * Harbour Project OR Free Software Foundation releases into a copy of
 * Harbour, as the General PUBLIC License permits, the exception does
 * NOT apply TO the code that you add IN this way.  TO avoid misleading
 * anyone as TO the status of such modified files, you must delete
 * this exception notice FROM them.
 *
 * IF you write modifications of your own FOR Harbour, it is your choice
 * whether TO permit this exception TO apply TO your modifications.
 * IF you DO NOT wish that, delete this exception notice.
 *
 */

#require "hbxlsxml"

#define pCANCELADA  "CANCELADA"
#define pAUTORIZADA "AUTORIZADA"

PROCEDURE Main()

   LOCAL oXml, oSheet, xarquivo := "example.xml"
   LOCAL xqtddoc, xttotnot, aDoc, nLinha, aNF
   LOCAL xEmpresa, xDataImp, xTitulo, xPeriodo, cStatusStyle
   
   REQUEST HB_CODEPAGE_PTISO

   Set( _SET_CODEPAGE, "PTISO" )
   hb_cdpSelect( "PTISO" )

   Set( _SET_DATEFORMAT, "yyyy-mm-dd" )

   oXml := ExcelWriterXML():New( xarquivo )
   oXml:setOverwriteFile( .T. )

   SetStyle( oXml )

   oSheet := oXml:addSheet( "Plan1" )

   oSheet:columnWidth(  1,  70 ) // N.Fiscal
   oSheet:columnWidth(  2,  70 ) // Data Emis.
   oSheet:columnWidth(  3,  80 ) // Situa��o
   oSheet:columnWidth(  4, 300 ) // Nome Cliente/Fornecedor
   oSheet:columnWidth(  5,  20 ) // UF
   oSheet:columnWidth(  6,  80 ) // Vlr.Tot.

   xEmpresa := "EMPRESA DEMONSTRA��O LTDA"
   xDataImp := DTOC(DATE())
   xTitulo := "RELAT�RIO PARA DEMONSTRAR XML EXCEL"
   xPeriodo := "Per�odo: "+DTOC(DATE()-49-40) + " a " + DTOC(DATE()-49-1)
   nLinha := 0

   oSheet:writeString( ++nLinha, 1, hb_StrToUTF8(xEmpresa) , "Cabec" )
   oSheet:cellMerge(     nLinha, 1, 4, 0 )
   oSheet:writeString( ++nLinha, 1, hb_StrToUTF8(xTitulo)  , "Cabec" )
   oSheet:cellMerge(     nLinha, 1, 4, 0 )
   oSheet:writeString( ++nLinha, 1, hb_StrToUTF8(xPeriodo) , "Cabec" )
   oSheet:cellMerge(     nLinha, 1, 4, 0 )
   oSheet:writeString( ++nLinha, 1,  hb_StrToUTF8("Data: " + xDataImp), "Cabec" )
   oSheet:cellMerge(     nLinha, 1, 4, 0 )

   oSheet:writeString( ++nLinha, 1, hb_StrToUTF8("N.Fiscal"          ), "textLeftBoldCor" )
   oSheet:writeString(   nLinha, 2, hb_StrToUTF8("Data Emiss�o"      ), "textLeftBoldCor" )
   oSheet:writeString(   nLinha, 3, hb_StrToUTF8("Situa��o"          ), "textLeftBoldCor" )
   oSheet:writeString(   nLinha, 4, hb_StrToUTF8("Cliente/Fornecedor"), "textLeftBoldCor" )
   oSheet:writeString(   nLinha, 5, hb_StrToUTF8("UF"                ), "textLeftBoldCor" )
   oSheet:writeString(   nLinha, 6, hb_StrToUTF8("Vlr.Tot."          ), "textRightBoldCor" )

   aDoc:= GetDocs()
   xqtddoc:= xttotnot:= 0

   FOR EACH aNF IN aDOC
      cStatusStyle:= IIF( aNF[ "situacao" ] == pCANCELADA, "Red", "" )
      nLinha++
      oSheet:writeString( nLinha, 1, hb_StrToUTF8( aNF[ "numeronf" ] ), "textLeft"+cStatusStyle )
      oSheet:writeString( nLinha, 2, hb_StrToUTF8( DTOC(aNF[ "dtemissao"]) ), "textLeft"+cStatusStyle )
      oSheet:writeString( nLinha, 3, hb_StrToUTF8( aNF[ "situacao" ] ), "textLeft"+cStatusStyle )
      oSheet:writeString( nLinha, 4, hb_StrToUTF8( aNF[ "nomecliente" ] ), "textLeft"+cStatusStyle )
      oSheet:writeString( nLinha, 5, hb_StrToUTF8( aNF[ "uf" ] ), "textLeft"+cStatusStyle )
      oSheet:writeNumber( nLinha, 6, aNF[ "valortotal" ], "numberRight"+cStatusStyle )
      xqtddoc++
      xttotnot += aNF[ "valortotal" ]
   NEXT

   oSheet:writeString( ++nLinha, 1, "", "textLeftBold" )
   oSheet:writeString(   nLinha, 2, "", "textLeftBold" )
   oSheet:writeString(   nLinha, 3, "", "textLeftBold" )
   oSheet:writeString(   nLinha, 4, "TOTAL ==> " + hb_ntos( xqtddoc ) + " documentos", "textLeftBold" )
   oSheet:writeString(   nLinha, 5, "", "textLeftBold" )
   oSheet:writeFormula( "Number", nLinha, 6, "=SUM(R[-40]C:R[-1]C)", "numberRightBold" )

   oXml:writeData( xarquivo )

   RETURN

 STATIC FUNCTION SetStyle( oXml )  
   LOCAL oObj
   oObj := oXml:addStyle( "textLeft" )
   oObj:alignHorizontal( "Left" )
   oObj:alignVertical( "Center" )
   oObj:setFontSize( 10 )

   oObj := oXml:addStyle( "textLeftRed" )
   oObj:alignHorizontal( "Left" )
   oObj:alignVertical( "Center" )
   oObj:setFontSize( 10 )
   oObj:bgColor( "red" )

   oObj := oXml:addStyle( "textLeftBold" )
   oObj:alignHorizontal( "Left" )
   oObj:alignVertical( "Center" )
   oObj:setFontSize( 10 )
   oObj:setFontBold()
   oObj:border( "Top", 2, "black", "Solid" )

   oObj := oXml:addStyle( "textLeftBoldCor" )
   oObj:alignHorizontal( "Left" )
   oObj:alignVertical( "Center" )
   oObj:setFontSize( 10 )
   oObj:setFontBold()
   oObj:bgColor( "lightblue" )
   oObj:alignWraptext()

   oObj := oXml:addStyle( "textRightBoldCor" )
   oObj:alignHorizontal( "Right" )
   oObj:alignVertical( "Center" )
   oObj:setFontSize( 10 )
   oObj:setFontBold()
   oObj:bgColor( "lightblue" )
   oObj:alignWraptext()

   oObj := oXml:addStyle( "numberRight" )
   oObj:alignHorizontal( "Right" )
   oObj:alignVertical( "Center" )
   oObj:setNumberFormat( "#,##0.00" )
   oObj:setFontSize( 10 )

   oObj := oXml:addStyle( "numberRightRed" )
   oObj:alignHorizontal( "Right" )
   oObj:alignVertical( "Center" )
   oObj:setNumberFormat( "#,##0.00" )
   oObj:setFontSize( 10 )
   oObj:bgColor( "red" )

   oObj := oXml:addStyle( "numberRightBold" )
   oObj:alignHorizontal( "Right" )
   oObj:alignVertical( "Center" )
   oObj:setNumberFormat( "#,##0.00" )
   oObj:setFontSize( 10 )
   oObj:setFontBold()
   oObj:border( "Top", 2, "black", "Solid" )

   oObj := oXml:addStyle( "Cabec" )
   oObj:alignHorizontal( "Left" )
   oObj:alignVertical( "Center" )
   oObj:setFontSize( 12 )
   oObj:setFontBold()

   oObj := oXml:addStyle( "CabecRight" )
   oObj:alignHorizontal( "Right" )
   oObj:alignVertical( "Center" )
   oObj:setFontSize( 12 )
   oObj:setFontBold()

   RETURN Nil

STATIC FUNCTION GetDocs()
   LOCAL aNames, aDoc:= { => }, i
   aNames:= GetNames()
   FOR i:= 1 TO 40
      aDoc[i]:= {;
         "numeronf"    => StrZero( i, 8 ),; 
         "dtemissao"   => Date() - 49 - i, ;
         "situacao"    => IIF( ((i-1) % 10) == 0, pCANCELADA, pAUTORIZADA ), ;
         "nomecliente" => aNames[i], ;
         "uf"          => "PR", ;
         "valortotal"  => i * 100 } 
   NEXT
   RETURN aDoc

STATIC FUNCTION GetNames()
   RETURN {;
     "SUPERMERCADO NA��ES LTDA                ",;
     "SINAI FRIOS COM�RCIO DE ALIMENTOS LTDA -",; 
     "SUPERMERCADO SANTO ANTONIO M.GUA�U LT   ",; 
     "NOVA MENDON�A - SUPERMERCADO LTDA       ",; 
     "MARANH�O SUPERMERCADOS S/A              ",; 
     "MARTINS COM�RCIO DE PRODUTOS ALIMENT�CIO",; 
     "SUPERMERCADO ESTRELA DE REGENTE FEIJ�   ",; 
     "SUPERMERCADO CASA ALIAN�A LTDA          ",; 
     "POEMA COM�RCIO G�NERO S ALIMENT�CIOS    ",; 
     "COMERCIAL KEYPAR REPRESENTA��ES E SUPERM",; 
     "SUPERMERCADO ALTA ROTA��O LTDA          ",; 
     "SUPERMERCADO SHIBATA TAUBAT� LTDA       ",; 
     "OP��O SUPERMERCADO DE S.B.C. LTDA - EPP ",; 
     "CASA AVENIDA COM�RCIO E IMPORTACAO LTDA ",; 
     "BEIRA RIO COM�RCIO EXPORTA��O E IMPORTA�",; 
     "RODRIGUES & PEREIRA CORDEIR�POLIS LTDA  ",; 
     "707 AUTO SERVI�O DE ALIMENTOS           ",; 
     "IGOMIC COM�RCIO DE ALIMENTOS LTDA       ",; 
     "EMP�RIO BOM GOSTO                       ",; 
     "CONTINENTAL COM�RCIO VAREJISTA LTDA.    ",; 
     "OP��O SUPERMERCADO                      ",; 
     "LUA AZUL - COM�RCIO DE PRODUTOS ALIMENT�",; 
     "SUPERMERCADO PEG PAG DOIS IRM�OS LTDA   ",; 
     "COM�RCIO DE ALIMENTOS ELION LTDA. - EPP ",; 
     "SUPERMERCADO CA�ULA LTDA                ",; 
     "MERCADO  S�O JOS�  DOMINGOS LTDA        ",; 
     "SUPERMERCADO SANTO ANTONIO M.GUA�U LTDA-",; 
     "DIPALMA COMERCIO DISTRIBUI��O E LOGISTIC",; 
     "SUPERMERCADO S�O JUDAS TADEU LTDA       ",; 
     "RIBEIRO & ALVES COM�RCIO DE ALIMENTOS VO",; 
     "SUPERMERCADO PEDROS�O LTDA              ",; 
     "LUZITA IND�STRIA E COM�RCIO LTDA        ",; 
     "MERCANTIL NOVA CURU�A LTDA              ",; 
     "IRM�OS MUFFATO & CIA LTDA.              ",; 
     "QUEIROZ - COMERCIO, ADMINISTRA��O & PLAN",; 
     "JOS� CARLOS MINATEL & CIA. LTDA. -EPP   ",; 
     "SUPERMERCADOS UNI�O SERV LTDA.          ",; 
     "SUPERMERCADO IRM�OS TEIXEIRA LTDA - EPP ",; 
     "COMERCIAL ZARAGOZA IMPORTA��O E EXPORTA�",; 
     "SUPERMERCADO CUCA DE ITANHA�M LTDA      " } 
