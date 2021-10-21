/**
 * @ts-nocheck
 * @OnlyCurrentDoc
 */
 
 /**
 * Tigger on open
 * @constructor
*/
function onOpen(e){
  myCreateBdColor();}

 /** 
 * Returns the Hexadecimal value of a cell's background color.
 * @param  {"A1"}    myRange "A1"or "A1:A20" or "A1:C20"
 * @param  {A1}      myTigger A1 or A1:A20 or A1:C20 on edit optional.
 * @return {#00ff00} myBackgroundRanges background color Hexadecimal.
 * @customfunction
 */
function myBackgroundRanges(myRange, myTigger) { 
  let myReturn;
  try{
    myReturn = SpreadsheetApp.getActiveSpreadsheet()
      .getActiveSheet()
        .getRange(myRange)
          .getBackgrounds();}  
  catch(myError) {
    myReturn = myMesaggeError(myError);}
  finally{
    return  myReturn;};}

/**
 * Returns the Rgb(Red, Green, Blue) value 0 to 255 of a cell's background color.
 * @constructor
 * @param  {"A1"}    myRange "A1"or "A1:A20" or "A1:C20"
 * @param  {A1}      myTigger A1 or A1:A20 or A1:C20 on edit optional.
 * @return {Rgb(255,255,255)} myAsRgbBackgroundRanges background color Rgb.
 * @customfunction
 */
function myAsRgbBackgroundRanges(myRange, myTigger) { 
  let myReturn, myReturnRgb;
  try{
    myReturn = SpreadsheetApp.getActiveSpreadsheet()
    .getActiveSheet()
      .getRange(myRange)
        .getBackgrounds();
    myReturnRgb = myReturn.map(row => row.map(cell => 'Rgb(' + parseInt(cell.slice(1,3),16) + ',' + parseInt(cell.slice(3,5),16) + ',' + parseInt(cell.slice(5,7),16) + ')'));}
  catch(myError) {
    myReturnRgb = myMesaggeError(myError);}
  finally {
    return myReturnRgb;}}

/**
 * Returns the name  cell's background color.
 * @constructor
 * @param  {"A1"}    myRange "A1"or "A1:A20" or "A1:C20"
 * @param  {A1}      myTigger A1 or A1:A20 or A1:C20 on edit optional.
 * @return {String} myNameBackgroundRanges background color name.
 * @customfunction
 */
function myNameBackgroundRanges(myRange, myTigger) { 
  let myHexColor = PropertiesService.getUserProperties();
  let myReturn, myReturnName;
  try{
    myReturn = SpreadsheetApp.getActiveSpreadsheet()
    .getActiveSheet()
      .getRange(myRange)
        .getBackgrounds();
    myReturnName = myReturn.map(row => row.map(cell => myHexColor.getProperty(cell)));}
  catch(myError) {
    myReturnName = myMesaggeError(myError);}
  finally {
    return myReturnName;}}
  
/**
* Return message error format string.
* @constructor
* @param  {'object'} myError - Object catch(myError) .
* @returns {String} message error complete.
*/      
function myMesaggeError(myError){
  let my= 'Name:' + myError.name + '\nMesagge:' + myError.message + '\nStack:'  + myError.stack;
  console.error(my);
  return my;}//fin function

/**
 * Create BD color table
 * @constructor
*/
function myCreateBdColor(){
  var myTableColors = PropertiesService.getUserProperties();
  if(myTableColors.getProperty('#8a2be2')){ console.info('BD Colors in stock');return;};
  console.info('Create BD Colors') 
var myBD = {
  '#8a2be2': 'Violeta Azul'
  ,'#ffebcd': 'Almendra blanqueada'
  ,'#f5f5f5': 'humo blanco'
  ,'#a52a2a': 'marrón'
  ,'#00ff7f': 'Primavera verde'
  ,'#daa520': 'Vara dorada'
  ,'#808000': 'Aceituna'
  ,'#808000': 'Aceituna'
  ,'#f5fffa': 'Acústico'
  ,'#7fffd4': 'Aguamarina'
  ,'#f0f8ff': 'Alice azul'
  ,'#ffff00': 'amarillo'
  ,'#ffffe0': 'Amarillo claro'
  ,'#ffd966': 'amarillo claro 1'
  ,'#ffe599': 'amarillo claro 2'
  ,'#fff2cc': 'amarillo claro 3'
  ,'#f1c232': 'amarillo oscuro 1'
  ,'#bf9000': 'amarillo oscuro 2'
  ,'#7f6000': 'amarillo oscuro 3'
  ,'#9acd32': 'Amarillo verde'
  ,'#9966cc': 'Amatista'
  ,'#cd5c5c': 'aned'
  ,'#fdf5e6': 'Antiguo'
  ,'#000080': 'Armada'
  ,'#4a86e8': 'azul aciano '
  ,'#0000ff': 'azul'
  ,'#1e90ff': 'Azul Dodger'
  ,'#4682b4': 'Azul acero'
  ,'#c9daf8': 'azul aciano 3'
  ,'#6d9eeb': 'Azul aciano claro 1'
  ,'#3c78d8': 'azul aciano oscuro 1'
  ,'#1155cc': 'azul aciano oscuro 2'
  ,'#1c4587': 'azul aciano oscuro 3'
  ,'#5f9ea0': 'Azul cadete'
  ,'#00bfff': 'Azul cielo azul'
  ,'#add8e6': 'Azul claro'
  ,'#6fa8dc': 'Azul claro 1'
  ,'#9fc5e8': 'Azul claro 2'
  ,'#cfe2f3': 'azul claro 3'
  ,'#b0c4de': 'Azul de acero ligero'
  ,'#6495ed': 'Azul de aciano'
  ,'#191970': 'Azul de medianoche'
  ,'#0000cd': 'Azul medio'
  ,'#00008b': 'Azul oscuro'
  ,'#3d85c6': 'azul oscuro 1'
  ,'#073763': 'azul oscuro 3'
  ,'#483d8b': 'Azul oscuro azul'
  ,'#b0e0e6': 'Azul pálido'
  ,'#7b68ee': 'Azul Pizarra Medio'
  ,'#4169e1': 'Azul real'
  ,'#f0ffff': 'Cian palido'
  ,'#fafad2': 'Barra de oro clara amarilla'
  ,'#b8860b': 'Barra dorada oscura'
  ,'#980000': 'Baya Roja'
  ,'#cc4125': 'baya roja clara 1'
  ,'#dd7e6b': 'baya roja clara 2'
  ,'#e6b8af': 'baya roja clara 3'
  ,'#a61c00': 'baya roja oscura 1'
  ,'#85200c': 'baya roja oscura 2'
  ,'#5b0f00': 'baya roja oscura 3'
  ,'#f5f5dc': 'Beige'
  ,'#ffffff': 'blanco'
  ,'#faebd7': 'Blanco antiguo'
  ,'#f8f8ff': 'Blanco fantasma'
  ,'#fffaf0': 'Blanco floral'
  ,'#228b22': 'Bosque verde'
  ,'#f0e68c': 'Caqui'
  ,'#0b5394': 'Caqui oscuro'
  ,'#bdb76b': 'Caqui oscuro'
  ,'#bdb76b': 'Caqui oscuro'
  ,'#d8bfd8': 'Cardo'
  ,'#dc143c': 'carmesí'
  ,'#d2691e': 'Chocolate'
  ,'#00ffff': 'cian'
  ,'#e0ffff': 'Cian claro'
  ,'#76a5af': 'cian claro 1'
  ,'#a2c4c9': 'cian claro 2'
  ,'#d0e0e3': 'cian claro 3'
  ,'#008b8b': 'Cian oscuro'
  ,'#45818e': 'cian oscuro 1'
  ,'#134f5c': 'cian oscuro 2'
  ,'#0c343d': 'cian oscuro 3'
  ,'#87ceeb': 'Cielo azul'
  ,'#87cefa': 'Cielo claro azul'
  ,'#dda0dd': 'Ciruela'
  ,'#93c47d': 'Claro verde 1'
  ,'#fff5ee': 'Concha'
  ,'#ff7f50': 'Coral'
  ,'#f08080': 'Coral ligero'
  ,'#dcdcdc': 'Ganancia boro'
  ,'#fffacd': 'Gasa de limón'
  ,'#f0fff0': 'Gotas de miel'
  ,'#800000': 'Granate'
  ,'#cccccc': 'Gris'
  ,'#808080': 'gris'
  ,'#d3d3d3': 'Gris claro'
  ,'#d9d9d9': 'gris claro 1'
  ,'#efefef': 'gris claro 2'
  ,'#f3f3f3': 'gris claro 3'
  ,'#a9a9a9': 'Gris oscuro'
  ,'#696969': 'Gris oscuro'
  ,'#b7b7b7': 'gris oscuro 1'
  ,'#999999': 'gris oscuro 2'
  ,'#666666': 'gris oscuro 3'
  ,'#434343': 'gris oscuro 4'
  ,'#708090': 'Gris pizarra'
  ,'#778899': 'Gris pizarra ligera'
  ,'#2f4f4f': 'Gris pizarra oscuro'
  ,'#4b0082': 'Índigo'
  ,'#b22222': 'Ladrillo de fuego'
  ,'#ffefd5': 'Látigo papaya'
  ,'#e6e6fa': 'Lavanda'
  ,'#fff0f5': 'Lavanda'
  ,'#faf0e6': 'Lino'
  ,'#a4c2f4': 'Luz azul acian 2'
  ,'#deb887': 'Madera'
  ,'#ff00ff': 'magenta'
  ,'#c27ba0': 'magenta claro 1'
  ,'#d5a6bd': 'magenta claro 2'
  ,'#ead1dc': 'magenta claro 3'
  ,'#8b008b': 'Magenta oscura'
  ,'#a64d79': 'magenta oscuro 1'
  ,'#741b47': 'magenta oscuro 2'
  ,'#4c1130': 'magenta oscuro 3'
  ,'#2e8b57': 'Mar verde'
  ,'#fffff0': 'Marfil'
  ,'#3cb371': 'Margen medio'
  ,'#f4a460': 'Marrón arenoso'
  ,'#8b4513': 'Marrón de montar'
  ,'#bc8f8f': 'Marrón rosado'
  ,'#66cdaa': 'Medio Aquamarine'
  ,'#ffe4e1': 'Misterosa'
  ,'#ffe4b5': 'Mocasín'
  ,'#7fff00': 'Monasterio'
  ,'#9370db': 'Morado medio'
  ,'#674ea7': 'morado oscuro 1'
  ,'#351c75': 'morado oscuro 2'
  ,'#ff9900': 'naranja'
  ,'#ffa500': 'naranja'
  ,'#f6b26b': 'naranja claro 1'
  ,'#f9cb9c': 'naranja claro 2'
  ,'#fce5cd': 'naranja claro 3'
  ,'#ff8c00': 'Naranja oscuro'
  ,'#e69138': 'naranja oscuro 1'
  ,'#b45f06': 'naranja oscuro 2'
  ,'#783f04': 'naranja oscuro 3'
  ,'#ffdead': 'Navajo blanco'
  ,'#000000': 'negro'
  ,'#fffafa': 'Nieve'
  ,'#ffd700': 'Oro'
  ,'#da70d6': 'Orquídea'
  ,'#ba55d3': 'Orquídea media'
  ,'#9932cc': 'Orquídea oscura'
  ,'#cd853f': 'Perú'
  ,'#6a5acd': 'Pizarra'
  ,'#c0c0c0': 'Plata'
  ,'#ffdab9': 'Puñalada'
  ,'#9900ff': 'púrpura'
  ,'#800080': 'Púrpura'
  ,'#800080': 'Púrpura'
  ,'#20124d': 'púrpura oscuro 3'
  ,'#ff0000': 'rojo'
  ,'#e06666': 'rojo claro 1'
  ,'#ea9999': 'rojo claro 2'
  ,'#f4cccc': 'rojo claro 3'
  ,'#ff4500': 'Rojo naranja'
  ,'#8b0000': 'Rojo oscuro'
  ,'#cc0000': 'rojo oscuro 1'
  ,'#990000': 'rojo oscuro 2'
  ,'#660000': 'rojo oscuro 3'
  ,'#db7093': 'Rojo violeta pálido'
  ,'#ff69b4': 'Rosa caliente'
  ,'#ffb6c1': 'Rosa claro'
  ,'#ff1493': 'Rosa profundo'
  ,'#ffc0cb': 'Rosado'
  ,'#fa8072': 'Salmón'
  ,'#ffa07a': 'Salmón ligero'
  ,'#e9967a': 'Salmón oscuro'
  ,'#fff8dc': 'Seda de maiz'
  ,'#ffe4c4': 'Sopa de mariscos'
  ,'#d2b48c': 'Tan'
  ,'#a0522d': 'Tierra de siena'
  ,'#ff6347': 'Tomate'
  ,'#f5deb3': 'Trigo'
  ,'#008080': 'Trullo'
  ,'#40e0d0': 'Turquesa'
  ,'#48d1cc': 'Turquesa mediana'
  ,'#00ced1': 'Turquesa oscura'
  ,'#afeeee': 'Turquesa pálida'
  ,'#eee8aa': 'Vara de oro pálido'
  ,'#00ff00': 'verde Luminoso'
  ,'#008000': 'Verde'
  ,'#adff2f': 'Verde amarillo'
  ,'#7cfc00': 'Verde césped'
  ,'#90ee90': 'Verde claro'
  ,'#b6d7a8': 'Verde claro 2'
  ,'#d9ead3': 'Verde claro 3'
  ,'#20b2aa': 'Verde claro verde'
  ,'#32cd32': 'Verde lima'
  ,'#6b8e23': 'Verde oliva'
  ,'#556b2f': 'Verde oliva verde'
  ,'#8fbc8f': 'Verde oscuro'
  ,'#006400': 'Verde oscuro'
  ,'#6aa84f': 'verde oscuro 1'
  ,'#38761d': 'verde oscuro 2'
  ,'#274e13': 'verde oscuro 3'
  ,'#98fb98': 'Verde pálido'
  ,'#00fa9a': 'Verde primavera verde'
  ,'#ee82ee': 'Violeta'
  ,'#8e7cc3': 'violeta claro 1'
  ,'#b4a7d6': 'violeta claro 2'
  ,'#d9d2e9': 'violeta claro 3'
  ,'#c71585': 'Violeta media'
  ,'#9400d3': 'Violeta oscuro'
  ,'#4285f4': 'Arándano'
  ,'#ea4335': 'Rojo brillante'
  ,'#fbbc04': 'Naranja vivido'
  ,'#34a853': 'Lima verde'
  ,'#ff6d01': 'Naranja oscuro 15'
  ,'#46bdc6': 'Azul 11'
  ,'#ff6d01': 'Naranja oscuro 15'
  ,'#ff6d01': 'Naranja oscuro 15'};
  myTableColors.setProperties(myBD);}

/**
 * Delete BD color table
 * @constructor
*/
function myDeleteBDColor(){
  console.info('Delete BD Colors');  
  PropertiesService.getUserProperties().deleteAllProperties();}
