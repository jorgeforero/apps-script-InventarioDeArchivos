/**
 * MyFiles
 * Obtener mis archivos en una hoja de cálculo y poder marcar archivos para
 * removerlos
 * Author: Jorge Forero - jorge.e.forero@gmail.com
 * 
 * WRNG: Actualizar ROOTFOLDERID para el borrado de los folder vacios
 */

// Id del folder root desde el cual se hace la búsqueda de los folderes vacios
const ROOTFOLDERID = '__ID_ROOT_FOLDER__'; 
const SHEETNAME = 'MyFIles';
const SHEETEMPTYF = 'emptyFolders';
const FOLDERURL = 'https://drive.google.com/drive/folders/';

/**
 * onOpen
 * Despliega el menú de con las opciones
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu( 'Opciones' )
      .addItem( '_ 🗄️ _ Obtener Mis Archivos', 'getMyFiles' )
      .addSeparator()
      .addItem( '_ ❌ _ Eliminar Archivos', 'removeMyFiles' )
      .addItem( '_ 🔀 _ Remover Marcados', 'moveRemoved' )
      .addItem( '_ 📁 _ Eliminar Folders Vacios', 'deleteEmptyFolders' )
      .addToUi();
};

// Basado en https://stackoverflow.com/questions/77838102/gdrive-search-all-files-in-mydrive-which-are-shared

/**
 * getMyFiles
 * Obtiene todos mis archivos y registra la información en una hoja de cálculo dada. 
 * La información que registra es: Name, Id, Type, Owner, Url 
 *   
 * @param {void} - void
 * @return {number} filesCounter - Número de archivos encontrados - Información en la hoja de cálculo dada ( si aplica )
 */
function getMyFiles() {
  const setLinkAsFormula = ( link, label ) => `=HYPERLINK("${link}", "${label}")`;
  // Cuenta activa
  let user = Session.getActiveUser().getEmail();
  let byte = 0.000001;
  let fileList = [];
  let pageToken;
  do {
    const obj = Drive.Files.list( { q: `'${user}' in owners and mimeType != 'application/vnd.google-apps.folder' and trashed=false`, 
      pageToken, pageSize: 1000, 
      fields: "nextPageToken,files( id, name, webViewLink, mimeType, size, createdTime, modifiedTime, parents )" } );
    // Si hay archivos los va adicionando a la lista de archivos
    if ( obj.files.length > 0 ) fileList = [ ...fileList, ...obj.files ];
    pageToken = obj.nextPageToken;
  } while ( pageToken );
  // Obtiene solo los datos requeridos del archivo
  const fileNames = fileList.map( ({ name, id, webViewLink, mimeType, size, createdTime, modifiedTime, parents }) => [ '', name, setLinkAsFormula( webViewLink, 'Ver' ), id, mimeType, ( size == null ) ? 0 : parseInt( size ) * byte, Utilities.formatDate( new Date( createdTime ), 'GMT-5', 'yyyy-MM-dd' ), Utilities.formatDate( new Date( modifiedTime ), 'GMT-5', 'yyyy-MM-dd' ), setLinkAsFormula( `${FOLDERURL}${parents}`, 'Folder' ) ] );
  // Registro de los datos de los archivos en la hoja de cálculo
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( SHEETNAME );
  // Limpia el contenido de la hoja
  sheet.getRange( 2, 1, sheet.getLastRow(), sheet.getLastColumn() ).clearContent();
  sheet.getRange( 2, 1, fileNames.length, fileNames[ 0 ].length ).setValues( fileNames );
  // Adiciona los check para poder marcar los archivos a borrar
  addChecks( sheet );
  // Retorna el número de archivo encontrados
  SpreadsheetApp.getActiveSpreadsheet().toast( `Se encontraron ${fileNames.length} archivos`, 'Status', 4 );
  return fileNames.length;
};

/**
 * addChecks
 * Adiciona los checks en la primera columna para poder marcar los archivos a borrar ( desde la segunda fila )
 * 
 * @param {object} sheet - Hoja de cálculo
 * @retunt {void} - checkboxes adicionados en la primera columna de la hoja dada 
 */
function addChecks( sheet ) {
  sheet.getRange( 2, 1, sheet.getLastRow(), 1 )
             .setBackground( '#b7e1cd' )
             .setVerticalAlignment( 'middle' ).setHorizontalAlignment( 'center' )
             .insertCheckboxes();
};

/**
 * removeMyFiles
 * A partir de la información obtenida por la función getFilesSharedWithMe en la hoja de calculo, se remueve
 * el permiso de editor de los archivos cuya columna RemoveMe y meEdit esten marcados en true.
 * La hoja calcula es actualizada en la columna RemoveMe indicando los archivos a los cuales le fue removido el permiso
 * 
 * @param {void} - void
 * @return {number} filesCounter - Número de archivos a los que se les removio el permiso de editor - Información actualizaza en la hoja ( si aplica )
 */
function removeMyFiles() {
  SpreadsheetApp.getActiveSpreadsheet().toast( `Working...`, 'Status', 4 );
  // Contadores
  let filesCounter = 0;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( SHEETNAME );
  // Obtiene los datos de los archivos encontrados por la función getFilesSharedWithMe
  let filesInfo = sheet.getDataRange().getValues();
  // Obtiene el header - Primera fila del arreglo
  let header = filesInfo.shift();
  // Obtiene los valores de la columna identificada con el nombre RemoveMe
  let flags = getColumnValues( header, filesInfo, 'RemoveMe' );
  for ( let indx=0; indx<filesInfo.length; indx++ ) {
    let record = getRowAsObject( filesInfo[ indx ], header );
    // Si la columna removeme esta vacia y la columna meedit es true, se remueve el permiso de editor
    if ( record.removeme == true ) {
      // Remueve el archivo sin dejarlo en la papelera
      res = Drive.Files.remove( record.id );
      filesCounter++;
      // Marca la celda del registro en proceso como removida
      flags[ indx ] = [ 'ReMoVeD' ];
    };    
  };//for
  // Si hubo cambios, se actualiza la hoja de cálculo fuente - Actualiza solo la columna RemoveMe para marcar los archivos removidos
  if ( filesCounter > 0 ) sheet.getRange( 2, getColumnIndex( header, 'RemoveMe' ) + 1, flags.length, 1 ).setValues( flags );
  SpreadsheetApp.getActiveSpreadsheet().toast( `Se removieron ${filesCounter} Archivos`, 'Status', 4 );
  return filesCounter;
};

/**
 * moveRemoved
 * Elimina los registros marcados con 'ReMoVed' en la hoja de files y los pasa a la 
 * hoja 'Removed'
 * 
 * @param {void} - void
 * @return {void} - Registro movidos a Removed ( Si los hay )
 */
function moveRemoved() {
  let book = SpreadsheetApp.getActive();
  let sheet = book.getSheetByName( SHEETNAME );
  let removedSheet = book.getSheetByName( SHEETEMPTYF );
  // Obtiene los datos 
  let files = sheet.getDataRange().getValues();
  let removedFiles = removedSheet.getDataRange().getValues();
  let deleted = 0;
  // Remueve los registros marcados como retirados en la hoja
  let indx = 1;
  while ( indx < files.length ) {
    let record = getRowAsObject( files[ indx ], files [ 0 ] );
    if ( record.removeme == 'ReMoVeD' ) {
      removedFiles.push( files[ indx ] );
      sheet.deleteRow( (indx + 1 ) - deleted );
      deleted++;
    };
    indx++;
  };
  // Si hubo registros encontrados, hace la actualización en las hojas
  if ( deleted > 0 ){
    removedSheet.clear( { contentsOnly: true } );
    removedSheet.getRange( 1, 1, removedFiles.length, removedFiles[ 0 ].length).setValues( removedFiles );
  };
  SpreadsheetApp.getActiveSpreadsheet().toast( `Se eliminaron ${ deleted } Archivos`, 'Status', 4 );
};

/*
 Tomado de https://stackoverflow.com/questions/49015740/delete-empty-folders-team-drive
 Borrado de folders vacios - WRNG: Requiere actualizar el ROOTFOLDERID para ejecutarlo
*/

/**
 * deleteEmptyFolders
 * Borra los folderes vacios
 * 
 * @param {void} - void
 * @return {void} - Folderes borrados registrados en la hoja emptyFolders
 */
function deleteEmptyFolders() {
  // Folder desde el cual se inicia la búsqueda  
  let parentFolder = DriveApp.getFolderById( ROOTFOLDERID );
  // Registro de los datos de los folders en emptyFolders
  let folders = parentFolder.getFolders();
  let  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( 'emptyFolders' );
  while ( folders.hasNext() ) {
    let childfolder = folders.next();
    recurseFolder( parentFolder, childfolder, sheet );
  };
};

/**
 * recurseFolder
 * Remueve los folderes vacios
 * 
 * @param {object} parentFolder - Objeto del folder padre 
 * @param {object} folder - Objeto al folder en proceso
 * @param {object} ss - Objeto a la hoja de cálculo donde se registran los folderes borrados
 * @return {void} - devuelve los valores en la recursión
 */
function recurseFolder( parentFolder, folder, ss ) {
  // Obtiene la lista de subfolders en el root de ese folder
  let childfolders = folder.getFolders();
  // Cuenta los subfolders
  while (childfolders.hasNext()) {
    let childfolder = childfolders.next();
    recurseFolder( folder, childfolder, ss );
  };
  // Obtiene la lista de archivos en el root de ese folder
  let hasFile = folder.getFiles().hasNext();
  // Obtiene la lista de folders en el root de ese folder
  let hasFolder = folder.getFolders().hasNext();
  // Si no hay folders y archivos, lo remueve
  if (!hasFile && !hasFolder) {
    // remueve el folder direct - requiere saber el parentFolder         
    // parentFolder.removeFolder(folder); 
    folder.setTrashed( true );
    // Registra el folder borrado
    ss.appendRow( [ folder.getId(), folder.getName()] );
  };
};

/**
 * getRowAsObject
 * Obtiene un objeto con los valores de la fila dada: RowData. Toma los nombres de las llaves del parámtero Header. Las llaves
 * son dadas en minusculas y los espacios reemplazados por _
 * 
 * @param {array} RowData - Arreglo con los datos de la fila de la hoja
 * @param {array} Header - Arreglo con los nombres del encabezado de la hoja
 * @return {object} obj - Objeto con los datos de la fila y las propiedades nombradas de acuerdo a Header
 */
 function getRowAsObject( RowData, Header ) {
  let obj = {};
  for ( let indx=0; indx<RowData.length; indx++ ) {
    obj[ Header[ indx ].toLowerCase().replace( /\s/g, '_' ) ] = RowData[ indx ];
  };//for
  return obj;
};

/**
 * getColumnValues
 * Obtiene todos los datos de la columna con nombre ColName
 * 
 * @param {string} ColName - Nombre de la columna de acuerdo a header (this.tbheader)
 * @return {array} - Arreglo con valores de la columna. Formato => [ [1],[2],[3] ]
 */
function getColumnValues( Header, Data, ColName ) {
  // Extrae la columna del arreglo Bidimensional a un arreglo lineal
  let colIndex = getColumnIndex( Header, ColName );
  return Data.map( function( value ) { return [ value[ colIndex ] ]; });
};

/**
 * getColumnIndex
 * Obtiene el indice (index-0) de la columna con el nombre Name
 * 
 * @param {string} Name - Nombre de la columna de acuerdo a el header
 * @return {integer} - Indice de Name en header o -1 sino lo encontró
 */
function getColumnIndex( Header, Name ) {
  return Header.indexOf( Name ); 
};
