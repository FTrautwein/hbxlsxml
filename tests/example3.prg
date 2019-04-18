#require "hbxlsxml"

PROCEDURE Main()

   LOCAL xml, sheet1, format4

   xml := ExcelWriterXML():New( "my file.xml" )

   sheet1 := xml:addSheet( "Plan 1" )

   format4 := xml:addStyle( "my style" )
   format4:setFontSize( 20 )
   format4:setFontColor( "yellow" )
   format4:bgColor( "blue" )

   sheet1:columnWidth( 1, 150 )
   sheet1:columnWidth( 2, 150 )
   sheet1:columnWidth( 3, 150 )

   sheet1:writeString( 1, 1, "celula 1_1", format4 )
#if 0
   sheet1:writeString( 1, 2, "celula 1_2", format4 )
#endif
   sheet1:writeString( 1, 3, "celula 1_3", format4 )
   sheet1:cellMerge( 1, 1, 1, 0 )

   sheet1:writeString( 2, 1, "celula 2_1", format4 )
   sheet1:writeString( 2, 2, "celula 2_2", format4 )
   sheet1:writeString( 2, 3, "celula 2_3", format4 )

   xml:writeData( "example3.xml" )

   RETURN
