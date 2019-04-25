/*
 * Copyright 2019 Fausto Di Creddo Trautwein, ftwein@yahoo.com.br
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

PROCEDURE Main()

   LOCAL oXml, oSheet, oObj, nLinha, xarquivo := "example4.xml"

   REQUEST HB_CODEPAGE_PTISO

   Set( _SET_CODEPAGE, "PTISO" )
   hb_cdpSelect( "PTISO" )

   Set( _SET_DATEFORMAT, "yyyy-mm-dd" )

   oXml := ExcelWriterXML():New( xarquivo )
   oXml:setOverwriteFile( .T. )

   oObj := oXml:addStyle( "textLeft" )
   oObj:alignHorizontal( "Left" )
   oObj:alignVertical( "Center" )
   oObj:setFontSize( 10 )

   oObj := oXml:addStyle( "textLeftSmall" )
   oObj:alignHorizontal( "Left" )
   oObj:alignVertical( "Center" )
   oObj:setFontSize( 5 )

   oObj := oXml:addStyle( "textLeftRed" )
   oObj:alignHorizontal( "Left" )
   oObj:alignVertical( "Center" )
   oObj:setFontSize( 10 )
   oObj:setFontBold()
   oObj:bgColor( "red" )
   oObj:alignWraptext()

   oObj := oXml:addStyle( "numberRight" )
   oObj:alignHorizontal( "Right" )
   oObj:alignVertical( "Center" )
   oObj:setNumberFormat( "#,##0.00;[Red]-#,##0.00;;@" )
   oObj:setFontSize( 10 )
   oObj:setFontBold()

   oSheet := oXml:addSheet( "Plan1" )

   oObj := oSheet
   oObj:columnWidth(  1, 100 ) 
   oObj:columnWidth(  2, 400 ) 
   oObj:columnWidth(  3,  70 ) 

   nLinha := 0

   oSheet:writeString( ++nLinha,  1, "Format", "textLeft" )
   oSheet:writeString(   nLinha,  2, "Text"  , "textLeft" )

   oSheet:writeString( ++nLinha,  1, "textLeft", "textLeft" )
   oSheet:writeString(   nLinha,  2, "Horizontal Left, Vertical Center, FontSize 10", "textLeft" )

   oSheet:writeString( ++nLinha,  1, "textLeftSmall", "textLeftSmall" )
   oSheet:writeString(   nLinha,  2, "Horizontal Left, Vertical Center, FontSize 5", "textLeftSmall" )

   oSheet:writeString( ++nLinha,  1, "textLeftRed", "textLeftRed" )
   oSheet:writeString(   nLinha,  2, "Horizontal Left, Vertical Center, FontSize 10, Color Red", "textLeftRed" )

   oSheet:writeString( ++nLinha,  1, "numberRight", "numberRight" )
   oSheet:writeNumber(   nLinha,  2, 1500.90, "numberRight" )

   oSheet:writeString( ++nLinha,  1, "numberRight", "numberRight" )
   oSheet:writeNumber(   nLinha,  2, -1500.90, "numberRight" )

   // BORDER
   oObj := oXml:addStyle( "borderAll" )
   oObj:alignHorizontal( "Center" )
   oObj:alignVertical( "Center" )
   oObj:fontSize( 10 )
   oObj:border( "All", 1, "blue", "Continuous" )

   oSheet:writeString( ++nLinha,  1, "borderAll", "borderAll" )
   oSheet:writeString(   nLinha,  2, "BORDER ALL, CONTINUOUS", "borderAll" )

   oObj := oXml:addStyle( "borderBottom" )
   oObj:alignHorizontal( "Center" )
   oObj:alignVertical( "Center" )
   oObj:fontSize( 10 )
   oObj:border( "Bottom", 2, "green", "Dash" )

   oSheet:writeString( ++nLinha,  1, "borderBottom", "borderBottom" )
   oSheet:writeString(   nLinha,  2, "BORDER BOTTOM, DASH", "borderBottom" )

   oXml:writeData( xarquivo )

   RETURN
