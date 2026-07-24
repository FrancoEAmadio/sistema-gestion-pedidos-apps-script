/*******************************************************
 * SISTEMA RDP - DIVISIÓN IV MATERIAL
 * Administración Central de Barrios Militares
 * VERSIÓN FINAL CONSOLIDADA
 *******************************************************/


/*******************************************************
 * CONFIGURACIÓN GENERAL
 *******************************************************/
const CONFIG = {

  ROOT_FOLDER_ID:
    '1RXONdXNMkB3y8aakwxPG5-bQeZfo222b',

  FORM_SHEET:
    'Form_Responses2',

  DB_SHEET:
    'Base de Datos RDP',

  TIMEZONE:
    'GMT-3',

  ADMIN_EMAIL:
    'relaciondepedidoacbmba@gmail.com',

  // ESTADOS
  STATUS_REVIEW:
    'EN REVISIÓN',

  STATUS_RECEIVED:
    'RECIBIDO',

  STATUS_REJECTED:
    'RECHAZADO',

  STATUS_SIGNED:
    'FIRMADO',

  STATUS_DEPOT:
    'EN DEPÓSITO DE SUMINISTROS',

  STATUS_PICKUP:
    'RETIRAR',

  STATUS_DELIVERED:
    'ENTREGADO',

  // LOGS
  LOG_RECEIVED:
    'MAIL RECIBIDO ENVIADO',

  LOG_REJECTED:
    'MAIL RECHAZADO ENVIADO',

  LOG_SIGNED:
    'ESTADO FIRMADO',

  LOG_DEPOT:
    'MAIL DEPÓSITO ENVIADO',

  LOG_PICKUP:
    'ESTADO RETIRAR',

  LOG_DELIVERED:
    'ESTADO ENTREGADO'
};


/*******************************************************
 * FORMULARIO
 *******************************************************/
function onFormSubmit(e) {

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const dbSheet =
      ss.getSheetByName(
        CONFIG.DB_SHEET
      );

    const values =
      e.values;

    // DATOS
    const timestamp =
      values[0];

    const pdfUrl =
      values[1];

    const dependencia =
      limpiarTexto(values[2]);

    const numeroRdp =
      limpiarTexto(values[3]);

    const unidadFuncional =
      limpiarTexto(values[4]);

    const categoria =
      limpiarTexto(values[5]);

    const mailUsuario =
      limpiarTexto(values[6]);

    // ARCHIVO
    const fileId =
      extraerId(pdfUrl);

    const originalFile =
      DriveApp.getFileById(fileId);

    // NOMBRE
    const nuevoNombre =
      construirNombreArchivo(
        dependencia,
        numeroRdp,
        unidadFuncional,
        categoria
      );

    originalFile.setName(
      nuevoNombre
    );

    // CARPETA DESTINO
    const folderDestino =
      obtenerCarpetaDependencia(
        dependencia
      );

    moverArchivo(
      originalFile,
      folderDestino
    );

    // DUPLICADO INFORMÁTICA
    if (
      categoria
        .toUpperCase()
        .trim() ===
      'INFORMATICA'
    ) {

      const root =
        DriveApp.getFolderById(
          CONFIG.ROOT_FOLDER_ID
        );

      const folders =
        root.getFoldersByName(
          'INFORMATICA'
        );

      if (
        folders.hasNext()
      ) {

        originalFile.makeCopy(
          nuevoNombre,
          folders.next()
        );
      }
    }

    // FILA
    const filaDestino =
      buscarPrimeraFilaLibre(
        dbSheet
      );

    // N° ORDEN
    dbSheet
      .getRange(filaDestino, 1)
      .setValue(
        filaDestino - 1
      );

    // FECHA INGRESO
    dbSheet
      .getRange(filaDestino, 2)
      .setValue(
        timestamp
      );

    // PDF
    const celdaLink =
      dbSheet.getRange(
        filaDestino,
        3
      );

    celdaLink.setRichTextValue(
      SpreadsheetApp
        .newRichTextValue()
        .setText(
          nuevoNombre
        )
        .setLinkUrl(
          originalFile.getUrl()
        )
        .build()
    );

    // DEPENDENCIA
    dbSheet
      .getRange(filaDestino, 4)
      .setValue(
        dependencia
      );

    // RDP
    dbSheet
      .getRange(filaDestino, 5)
      .setValue(
        numeroRdp
      );

    // UNIDAD
    dbSheet
      .getRange(filaDestino, 6)
      .setValue(
        unidadFuncional
      );

    // CATEGORÍA
    dbSheet
      .getRange(filaDestino, 7)
      .setValue(
        categoria
      );

    // ESTADO
    dbSheet
      .getRange(filaDestino, 8)
      .setValue(
        CONFIG.STATUS_REVIEW
      );

    // FECHA ESTADO
    dbSheet
      .getRange(filaDestino, 9)
      .setValue(
        obtenerFechaActual()
      );

    // MAIL
    dbSheet
      .getRange(filaDestino, 10)
      .setValue(
        mailUsuario
      );

    // LOG VACÍO
    dbSheet
      .getRange(filaDestino, 12)
      .setValue('');

  } catch (error) {

    Logger.log(error);

    MailApp.sendEmail({

      to:
        CONFIG.ADMIN_EMAIL,

      subject:
        'ERROR SISTEMA RDP - onFormSubmit',

      htmlBody: `
        <h2>Error en onFormSubmit</h2>
        <pre>${error}</pre>
      `
    });
  }
}


/*******************************************************
 * EDICIÓN
 *******************************************************/
function onEdit(e) {

  try {

    const sheet =
      e.range.getSheet();

    if (
      sheet.getName() !==
      CONFIG.DB_SHEET
    ) return;

    const row =
      e.range.getRow();

    const column =
      e.range.getColumn();

    if (column !== 8) return;

    const nuevoEstado =
      e.range.getValue();

    const estadoAnterior =
      e.oldValue || '';

    // RECUPERACIÓN DE FILAS BORRADAS
    const esRecuperacion =
      estadoAnterior === '' &&
      nuevoEstado ===
      CONFIG.STATUS_REVIEW;

    // VALIDAR FLUJO
    if (
      !validarCambioEstado(
        estadoAnterior,
        nuevoEstado
      )
    ) {

      SpreadsheetApp
        .getUi()
        .alert(
          'Cambio de estado inválido.'
        );

      e.range.setValue(
        estadoAnterior
      );

      return;
    }

    // DATOS
    const dependencia =
      sheet.getRange(row, 4)
        .getValue();

    const numeroRdp =
      sheet.getRange(row, 5)
        .getValue();

    const unidadFuncional =
      sheet.getRange(row, 6)
        .getValue();

    const categoria =
      sheet.getRange(row, 7)
        .getValue();

    const mailUsuario =
      sheet.getRange(row, 10)
        .getValue();

    const observaciones =
  sheet.getRange(row, 11)
    .getValue();

    const logs =
      sheet.getRange(row, 12)
        .getValue();

    // RECHAZADO SIN OBS
    if (
      nuevoEstado ===
      CONFIG.STATUS_REJECTED
    ) {

      if (
        !observaciones ||
        observaciones.trim() === ''
      ) {

        SpreadsheetApp
          .getUi()
          .alert(
            'Debe ingresar observaciones.'
          );

        e.range.setValue(
          estadoAnterior
        );

        return;
      }
    }

    // FECHA ESTADO
    sheet
      .getRange(row, 9)
      .setValue(
        obtenerFechaActual()
      );

    // SI ES RECUPERACIÓN
    // NO HACER MAILS NI LOGS
    if (esRecuperacion){
      return;
    }

    /***************************************************
     * RECIBIDO
     ***************************************************/
    if (
      nuevoEstado ===
      CONFIG.STATUS_RECEIVED
      ) {

      if (
        !logs.includes(
          CONFIG.LOG_RECEIVED
        )
      ) {

        enviarMailRecibido(
          mailUsuario,
          dependencia,
          numeroRdp,
          categoria,
          unidadFuncional
        );

        agregarLog(
          sheet,
          row,
          CONFIG.LOG_RECEIVED
        );
      }
    }

    /***************************************************
     * RECHAZADO
     ***************************************************/

    if (
      nuevoEstado ===
      CONFIG.STATUS_REJECTED
    ) {

      if (
        !logs.includes(
          CONFIG.LOG_REJECTED
        )
      ) {

        enviarMailRechazado(
          mailUsuario,
          dependencia,
          numeroRdp,
          categoria,
          unidadFuncional,
          observaciones
        );

        moverARechazadas(
          sheet,
          row,
          categoria
        );

        agregarLog(
          sheet,
          row,
          CONFIG.LOG_REJECTED
        );
      }
    }

    /***************************************************
     * FIRMADO
     ***************************************************/
    if (
      nuevoEstado ===
      CONFIG.STATUS_SIGNED
    ) {

      if (
        !logs.includes(
          CONFIG.LOG_SIGNED
        )
      ) {

        agregarLog(
          sheet,
          row,
          CONFIG.LOG_SIGNED
        );
      }
    }

    /***************************************************
     * DEPÓSITO
     ***************************************************/
    if (
      nuevoEstado ===
      CONFIG.STATUS_DEPOT
    ) {

      if (
        !logs.includes(
          CONFIG.LOG_DEPOT
        )
      ) {

        enviarMailDeposito(
          mailUsuario,
          dependencia,
          numeroRdp,
          categoria,
          unidadFuncional
        );

        agregarLog(
          sheet,
          row,
          CONFIG.LOG_DEPOT
        );
      }
    }

    /***************************************************
     * RETIRAR
     ***************************************************/
    if (
      nuevoEstado ===
      CONFIG.STATUS_PICKUP
    ) {

      if (
        !logs.includes(
          CONFIG.LOG_PICKUP
        )
      ) {

        agregarLog(
          sheet,
          row,
          CONFIG.LOG_PICKUP
        );
      }
    }

    /***************************************************
     * ENTREGADO
     ***************************************************/
    if (
      nuevoEstado ===
      CONFIG.STATUS_DELIVERED
    ) {

      if (
        !logs.includes(
          CONFIG.LOG_DELIVERED
        )
      ) {

        agregarLog(
          sheet,
          row,
          CONFIG.LOG_DELIVERED
        );
      }
    }

  } catch (error) {

    Logger.log(error);

    MailApp.sendEmail({

      to:
        CONFIG.ADMIN_EMAIL,

      subject:
        'ERROR SISTEMA RDP - onEdit',

      htmlBody: `
        <h2>Error en onEdit</h2>
        <pre>${error}</pre>
      `
    });
  }
}


/*******************************************************
 * VALIDAR FLUJO
 *******************************************************/

function validarCambioEstado(
  anterior,
  nuevo
) {

  const flujo = {

    '':[
      'EN REVISIÓN'
    ],

    'EN REVISIÓN':
      ['RECIBIDO'],

    'RECIBIDO':
      [
        'RECHAZADO',
        'FIRMADO'
      ],

    'FIRMADO':
      [
       'EN DEPÓSITO DE SUMINISTROS'
      ],

    'EN DEPÓSITO DE SUMINISTROS':
      ['RETIRAR'],

    'RETIRAR':
      ['ENTREGADO']
  };

  if (!flujo[anterior]) {
    return false;
  }

  return flujo[anterior]
    .includes(nuevo);
}

/*******************************************************
 * CARPETAS
 *******************************************************/
function obtenerCarpetaDependencia(
  dependencia
) {

  const root =
    DriveApp.getFolderById(
      CONFIG.ROOT_FOLDER_ID
    );

  const bbmm =
    [
      'BMGB',
      'BMMA',
      'BMVM',
      'BMGSM',
      'BMSC'
    ];

  let carpetaPadre;

  if (
    bbmm.includes(
      dependencia
    )
  ) {

    carpetaPadre =
      root
        .getFoldersByName(
          'BBMM'
        )
        .next();

  } else {

    carpetaPadre =
      root
        .getFoldersByName(
          'DIVS'
        )
        .next();
  }

  const folders =
    carpetaPadre
      .getFoldersByName(
        dependencia
      );

  if (
    !folders.hasNext()
  ) {

    throw new Error(
      `No existe carpeta: ${dependencia}`
    );
  }

  return folders.next();
}

/*******************************************************
 * MOVER ARCHIVO
 *******************************************************/

function moverArchivo(
  file,
  folderDestino
) {

  const padres =
    file.getParents();

  let carpetaOriginalId =
    null;

  if (
    padres.hasNext()
  ) {

    carpetaOriginalId =
      padres.next().getId();
  }

  Drive.Files.update(
    {},
    file.getId(),
    null,
    {
      addParents:
        folderDestino.getId(),

      removeParents:
        carpetaOriginalId
    }
  );
}

/*******************************************************
 * MOVER A RECHAZADAS
 *******************************************************/
function moverARechazadas(
  sheet,
  row,
  categoria
){

  const rich =
    sheet
      .getRange(row,3)
      .getRichTextValue();

  if(!rich) return;

  const url =
    rich.getLinkUrl();

  if(!url) return;

  const fileId =
    extraerId(url);

  const archivo =
    DriveApp.getFileById(
      fileId
    );

  // MOVER ORIGINAL
  const root =
    DriveApp.getFolderById(
      CONFIG.ROOT_FOLDER_ID
    );

  const carpetasRechazadas =
    root.getFoldersByName(
      'RECHAZADAS'
    );

  if(
    !carpetasRechazadas.hasNext()
  ){
    throw new Error(
      'No existe carpeta RECHAZADAS'
    );
  }

const rechazadas =
  carpetasRechazadas.next();

  moverArchivo(
    archivo,
    rechazadas
  );

  // SI ES INFORMÁTICA
  if(
    categoria
      .toUpperCase()
      .trim() ===
      'INFORMATICA'
  ){

    const carpetaInfo =
      root
      .getFoldersByName(
        'INFORMATICA'
      );

    if(
      carpetaInfo.hasNext()
    ){

      const folder =
        carpetaInfo.next();

      const archivos =
        folder.getFilesByName(
          archivo.getName()
        );

      while(
        archivos.hasNext()
      ){

        archivos
          .next()
          .setTrashed(true);
      }
    }
  }
}

/*******************************************************
 * RDP COMPLETA
 *******************************************************/
 
function construirRdpCompleta(
  dependencia,
  numeroRdp,
  categoria,
  unidadFuncional
) {

  const anio =
    new Date().getFullYear();

  unidadFuncional =
    limpiarTexto(
      unidadFuncional
    );

  dependencia =
    limpiarTexto(
      dependencia
    );

  numeroRdp =
    limpiarTexto(
      numeroRdp
    );

  categoria =
    limpiarTexto(
      categoria
    );

  let rdp =
    `${dependencia} - ${numeroRdp}/${anio}`;

  if (
    unidadFuncional !== ''
  ) {

    rdp +=
      ` - ${unidadFuncional}`;
  }

  rdp +=
    ` - ${categoria}`;

  return rdp;
}

/*******************************************************
 * NOMBRE ARCHIVO
 *******************************************************/
function construirNombreArchivo(
  dependencia,
  numeroRdp,
  unidadFuncional,
  categoria
) {

  return construirRdpCompleta(
    dependencia,
    numeroRdp,
    categoria,
    unidadFuncional
  );
}


/*******************************************************
 * MAIL RECIBIDO
 *******************************************************/
function enviarMailRecibido(
  mail,
  dependencia,
  numeroRdp,
  categoria,
  unidadFuncional
) {

  const rdp =
    construirRdpCompleta(
      dependencia,
      numeroRdp,
      categoria,
      unidadFuncional
    );

  MailApp.sendEmail({

    to: mail,

    subject:
      `${rdp} (RECIBIDO)`,

    htmlBody: `
      <p>
      Su Relación de Pedido fue recibida correctamente y se encuentra en proceso de control por parte del Jefe División IV Material.
      </p>

      <br>

      <b>RDP:</b> ${rdp}

      <br><br>

      Div IV Mat - ACBMBA
    `
  });
}


/*******************************************************
 * MAIL RECHAZADO
 *******************************************************/
function enviarMailRechazado(
  mail,
  dependencia,
  numeroRdp,
  categoria,
  unidadFuncional,
  motivo
) {

  const rdp =
    construirRdpCompleta(
      dependencia,
      numeroRdp,
      categoria,
      unidadFuncional
    );

  MailApp.sendEmail({

    to: mail,

    subject:
      `${rdp} (RECHAZADO)`,

    htmlBody: `
      <p>
      Su Relación de Pedido fue rechazada por la siguiente observación:
      </p>

      <blockquote>
      ${motivo}
      </blockquote>

      <br>

      <b>RDP:</b> ${rdp}

      <br><br>

      Div IV Mat - ACBMBA
    `
  });
}


/*******************************************************
 * MAIL DEPÓSITO
 *******************************************************/
function enviarMailDeposito(
  mail,
  dependencia,
  numeroRdp,
  categoria,
  unidadFuncional
) {

  const rdp =
    construirRdpCompleta(
      dependencia,
      numeroRdp,
      categoria,
      unidadFuncional
    );

  MailApp.sendEmail({

    to: mail,

    subject:
      `${rdp} (EN DEPÓSITO DE SUMINISTROS)`,

    htmlBody: `
      <p>
      Su Relación de Pedido ya fue firmada por el Jefe División IV Material y remitida al Depósito de Suministros para su preparación.
      </p>

      <br>

      <p>
      Para coordinar el retiro de los efectos, comuníquese directamente con el Encargado del Depósito de Suministros.
      </p>

      <br>

      <b>RDP:</b> ${rdp}

      <br><br>

      Div IV Mat - ACBMBA
    `
  });
}


/*******************************************************
 * LOG
 *******************************************************/
function agregarLog(
  sheet,
  row,
  texto
) {

  const fecha =
    obtenerFechaActual();

  const actual =
    sheet
      .getRange(row, 12)
      .getValue();

  // EVITAR DUPLICADOS
  if (
    actual.includes(texto)
  ) {
    return;
  }

  const nuevo =
    actual
      ? actual +
        '\n' +
        texto +
        ' - ' +
        fecha
      : texto +
        ' - ' +
        fecha;

  sheet
    .getRange(row, 12)
    .setValue(
      nuevo
    );
}


/*******************************************************
 * FECHA
 *******************************************************/
function obtenerFechaActual() {

  return Utilities.formatDate(
    new Date(),
    CONFIG.TIMEZONE,
    'dd/MM/yyyy HH:mm'
  );
}


/*******************************************************
 * EXTRAER ID
 *******************************************************/
function extraerId(url) {

  const match =
    url.match(/[-\w]{25,}/);

  if (!match) {

    throw new Error(
      'No se pudo extraer ID.'
    );
  }

  return match[0];
}


/*******************************************************
 * LIMPIAR TEXTO
 *******************************************************/
function limpiarTexto(texto) {

  if (!texto) return '';

  return texto
    .toString()
    .trim();
}


/*******************************************************
 * BUSCAR FILA LIBRE
 *******************************************************/
function buscarPrimeraFilaLibre(
  sheet
) {

  const ultimaFila =
    sheet.getLastRow();

  // HOJA VACÍA
  if (ultimaFila < 2) {
    return 2;
  }

  const datos =
    sheet
      .getRange(
        2,
        4,
        ultimaFila - 1,
        1
      )
      .getValues();

  for (
    let i = 0;
    i < datos.length;
    i++
  ) {

    if (!datos[i][0]) {

      return i + 2;
    }
  }

  return ultimaFila + 1;
}