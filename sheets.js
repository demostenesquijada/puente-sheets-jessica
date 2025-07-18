const { google } = require('googleapis');

// ===========================
// BLOQUE 1 – SHEET Y HOJAS
// ===========================

async function crearSheet(serviceAccount, titulo) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const res = await sheets.spreadsheets.create({
      resource: {
        properties: { title: titulo }
      }
    });
    return res.data.spreadsheetId;
  } catch (error) {
    throw new Error(`Error al crear el Sheet: ${error.message}`);
  }
}

async function crearHoja(serviceAccount, spreadsheetId, nombreHoja) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: {
        requests: [
          {
            addSheet: {
              properties: { title: nombreHoja }
            }
          }
        ]
      }
    });
    return true;
  } catch (error) {
    throw new Error(`Error al crear la hoja: ${error.message}`);
  }
}

async function buscarHoja(serviceAccount, spreadsheetId, nombreHoja) {
  console.log(`🔍 Buscando hoja "${nombreHoja}" en Spreadsheet ${spreadsheetId}`);
  const auth = new google.auth.GoogleAuth({
    credentials: serviceAccount,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });

  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  const metadata = await sheets.spreadsheets.get({
    spreadsheetId,
  });
  const disponibles = metadata.data.sheets.map(s => s.properties.title).join(', ');
  console.log(`📂 Hojas disponibles en el documento: ${disponibles}`);

  const hojaEncontrada = metadata.data.sheets.find(
    (sheet) =>
      sheet.properties.title.trim().toLowerCase() === nombreHoja.trim().toLowerCase()
  );
  console.log(hojaEncontrada ? '✅ Hoja encontrada' : '❌ Hoja NO encontrada');
  return hojaEncontrada ? true : false;
}

async function leerHoja(serviceAccount, spreadsheetId, sheetName) {
  const auth = new google.auth.GoogleAuth({
    credentials: serviceAccount,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });

  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A1:Z1000`, // puedes ajustar el rango si es necesario
  });

  return response.data.values;
}

async function obtenerSheet(serviceAccount, spreadsheetId) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const res = await sheets.spreadsheets.get({
      spreadsheetId
    });
    return res.data;
  } catch (error) {
    throw new Error(`Error al obtener metadatos del Sheet: ${error.message}`);
  }
}

async function obtenerHoja(serviceAccount, spreadsheetId, nombreHoja) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const metadata = await sheets.spreadsheets.get({
      spreadsheetId
    });
    const hoja = metadata.data.sheets.find(
      (sheet) => sheet.properties.title === nombreHoja
    );
    if (!hoja) throw new Error('Hoja no encontrada');
    return hoja;
  } catch (error) {
    throw new Error(`Error al obtener metadatos de la hoja: ${error.message}`);
  }
}

// ===========================
// BLOQUE 2 – FILAS
// ===========================

async function crearFila(serviceAccount, spreadsheetId, sheetName, datos) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    const fila = [
      datos.FECHA,
      datos.EJE,
      datos.PROYECTO,
      datos.SUBPROYECTO,
      datos.TEMA,
      datos.DETALLE,
      datos.ESTADO,
      datos.PRIORIDAD,
      datos.ASIGNADO_A,
      datos.OBSERVACIONES,
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetName}`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [fila]
      }
    });
    return true;
  } catch (error) {
    throw new Error(`Error al crear fila: ${error.message}`);
  }
}

async function leerFilas(serviceAccount, spreadsheetId, sheetIdOrName) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetIdOrName}`
    });
    return res.data.values || [];
  } catch (error) {
    throw new Error(`Error al leer filas: ${error.message}`);
  }
}

async function buscarFila(serviceAccount, spreadsheetId, sheetName, columna, valorBuscado) {
  // Alias simplificado: delega en buscarFilasPorCondiciones
  const condiciones = [{ campo: columna, valor: valorBuscado }];
  const resultados = await buscarFilasPorCondiciones(serviceAccount, spreadsheetId, sheetName, condiciones);
  // Solo retorna la primera coincidencia para mantener comportamiento previo
  return resultados.length > 0 ? resultados[0] : null;
}

// ===========================
// buscarFilasPorCondiciones
// ===========================
async function buscarFilasPorCondiciones(serviceAccount, spreadsheetId, sheetName, condiciones) {
  try {
    const filas = await leerFilas(serviceAccount, spreadsheetId, sheetName);
    if (!filas || filas.length === 0) return [];

    const encabezados = filas[0];

    // Validar que todas las columnas existan
    for (const condicion of condiciones) {
      if (!encabezados.includes(condicion.campo)) {
        throw new Error(`Columna no encontrada: ${condicion.campo}`);
      }
    }

    // Filtrar filas que cumplan todas las condiciones
    const resultados = filas
      .slice(1) // Omitir encabezados
      .filter(fila => {
        return condiciones.every(condicion => {
          const idx = encabezados.indexOf(condicion.campo);
          const valorCelda = (fila[idx] || '').toString().trim();
          return valorCelda === condicion.valor;
        });
      })
      .map(fila => {
        return encabezados.reduce((acc, key, i) => {
          acc[key] = fila[i] || '';
          return acc;
        }, {});
      });

    return resultados;
  } catch (error) {
    throw new Error(`Error al buscar filas por condiciones: ${error.message}`);
  }
}




// ===========================
// leerFila (personalizado)
// ===========================
// ===========================
// leerFila (ajustado: index 1 = primera fila de datos, encabezado en 0)
// ===========================
async function leerFila(serviceAccount, spreadsheetId, sheetName, opciones) {
  try {
    const filas = await leerFilas(serviceAccount, spreadsheetId, sheetName);
    if (!filas || filas.length === 0) throw new Error('La hoja está vacía');

    if (opciones.index !== undefined) {
      const idxUsuario = parseInt(opciones.index, 10);
      // Ahora: index=1 es la primera fila de datos (fila 0 = encabezado)
      if (isNaN(idxUsuario) || idxUsuario <= 0 || idxUsuario > filas.length - 1) {
        throw new Error('Índice fuera de rango o inválido (fila 1 = primera fila de datos)');
      }
      const idxReal = idxUsuario; // fila 1 = primera fila de datos, encabezado sigue en 0
      const encabezados = filas[0];
      const datos = filas[idxReal];
      const objeto = encabezados.reduce((acc, key, i) => {
        acc[key] = datos[i] || '';
        return acc;
      }, {});
      return objeto;
    }

    if (opciones.campo && opciones.valor !== undefined) {
      const encabezados = filas[0];
      const colIdx = encabezados.indexOf(opciones.campo);
      if (colIdx === -1) throw new Error('Columna no encontrada');

      const fila = filas.find((fila, i) => i > 0 && (fila[colIdx] || '').toString().trim() === opciones.valor);
      if (!fila) throw new Error('No se encontró ninguna fila que coincida');
      const objeto = encabezados.reduce((acc, key, i) => {
        acc[key] = fila[i] || '';
        return acc;
      }, {});
      return objeto;
    }

    throw new Error('Debes especificar index o campo + valor');
  } catch (error) {
    throw new Error(`Error al leer fila: ${error.message}`);
  }
}

// ===========================
// leerFilasPorIndices (ajustado: indices de usuario 1 = primera fila de datos)
// ===========================
async function leerFilasPorIndices(serviceAccount, spreadsheetId, sheetName, indices) {
  try {
    const filas = await leerFilas(serviceAccount, spreadsheetId, sheetName);
    if (!filas || filas.length === 0) throw new Error('La hoja está vacía');

    const encabezados = filas[0];

    const resultados = indices.map(idxUsuario => {
      const idxReal = parseInt(idxUsuario, 10);
      // Usuario pasa 1 = primera fila de datos, así que accedemos a filas[idxReal]
      if (isNaN(idxReal) || idxReal <= 0 || idxReal > filas.length - 1) {
        return { error: `Índice ${idxUsuario} fuera de rango o inválido (fila 1 = primera fila de datos)` };
      }
      const datos = filas[idxReal];
      return encabezados.reduce((acc, key, i) => {
        acc[key] = datos[i] || '';
        return acc;
      }, {});
    });

    return resultados;
  } catch (error) {
    throw new Error(`Error al leer filas por índices: ${error.message}`);
  }
}

// ===========================
// borrarFila (ajustado: usuario pasa 1 = primera fila de datos, sumar 1 para la API)
// ===========================
async function borrarFila(serviceAccount, spreadsheetId, sheetName, numeroFila) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    // Ajuste: sumar 1 para considerar encabezado en API
    const filaReal = numeroFila + 1;
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const hoja = meta.data.sheets.find(s => s.properties.title === sheetName);
    if (!hoja) throw new Error('Hoja no encontrada');
    const sheetId = hoja.properties.sheetId;
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex: filaReal - 1,
              endIndex: filaReal
            }
          }
        }]
      }
    });
    return true;
  } catch (error) {
    throw new Error(`Error al borrar fila: ${error.message}`);
  }
}

// ===========================
// BLOQUE 3 – COLUMNAS
// ===========================

async function crearColumna(serviceAccount, spreadsheetId, sheetName, tituloColumna) {
  try {
    // Leer los datos existentes
    const filas = await leerFilas(serviceAccount, spreadsheetId, sheetName);
    // Añadir el título al final de la primera fila (encabezados)
    if (!filas.length) throw new Error('La hoja no tiene filas para agregar columna');
    filas[0].push(tituloColumna);
    // Añadir celdas vacías a cada fila subsiguiente
    for (let i = 1; i < filas.length; i++) {
      filas[i].push('');
    }
    // Sobrescribir el rango completo
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: filas }
    });
    return true;
  } catch (error) {
    throw new Error(`Error al crear columna: ${error.message}`);
  }
}

// Devuelve solo el array de encabezados (primera fila de la hoja)
async function leerEncabezado(serviceAccount, spreadsheetId, sheetName) {
  try {
    const filas = await leerFilas(serviceAccount, spreadsheetId, sheetName);
    if (!filas.length) return [];
    return filas[0];
  } catch (error) {
    throw new Error(`Error al leer encabezados: ${error.message}`);
  }
}

async function buscarColumna(serviceAccount, spreadsheetId, sheetName, nombreColumna) {
  try {
    const filas = await leerFilas(serviceAccount, spreadsheetId, sheetName);
    if (!filas.length) return -1;
    const encabezados = filas[0];
    return encabezados.indexOf(nombreColumna);
  } catch (error) {
    throw new Error(`Error al buscar columna: ${error.message}`);
  }
}

async function obtenerColumna(serviceAccount, spreadsheetId, sheetName, nombreColumna) {
  try {
    const filas = await leerFilas(serviceAccount, spreadsheetId, sheetName);
    if (!filas.length) return [];
    const encabezados = filas[0];
    const idx = encabezados.indexOf(nombreColumna);
    if (idx === -1) throw new Error('Columna no encontrada');
    const datos = [];
    for (let i = 1; i < filas.length; i++) {
      datos.push(filas[i][idx]);
    }
    return datos;
  } catch (error) {
    throw new Error(`Error al obtener columna: ${error.message}`);
  }
}

async function borrarColumna(serviceAccount, spreadsheetId, sheetName, nombreColumna) {
  try {
    // Obtener sheetId y el índice de la columna
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const hoja = meta.data.sheets.find(s => s.properties.title === sheetName);
    if (!hoja) throw new Error('Hoja no encontrada');
    const sheetId = hoja.properties.sheetId;
    // Obtener índice de columna
    const filas = await leerFilas(serviceAccount, spreadsheetId, sheetName);
    if (!filas.length) throw new Error('La hoja está vacía');
    const encabezados = filas[0];
    const colIdx = encabezados.indexOf(nombreColumna);
    if (colIdx === -1) throw new Error('Columna no encontrada');
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId,
              dimension: 'COLUMNS',
              startIndex: colIdx,
              endIndex: colIdx + 1
            }
          }
        }]
      }
    });
    return true;
  } catch (error) {
    throw new Error(`Error al borrar columna: ${error.message}`);
  }
}

// ===========================
// BLOQUE 4 – CELDAS
// ===========================

async function llenarCelda(serviceAccount, spreadsheetId, sheetName, celda, valor) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!${celda}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [[valor]] }
    });
    return true;
  } catch (error) {
    throw new Error(`Error al llenar celda: ${error.message}`);
  }
}

async function leerCelda(serviceAccount, spreadsheetId, sheetName, celda) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!${celda}`
    });
    return res.data.values && res.data.values[0] ? res.data.values[0][0] : null;
  } catch (error) {
    throw new Error(`Error al leer celda: ${error.message}`);
  }
}

async function buscarCelda(serviceAccount, spreadsheetId, sheetName, valorBuscado) {
  try {
    const filas = await leerFilas(serviceAccount, spreadsheetId, sheetName);
    const coincidencias = [];

    for (let i = 0; i < filas.length; i++) {
      for (let j = 0; j < filas[i].length; j++) {
        if (filas[i][j] === valorBuscado) {
          // Convertir columna numérica a letra
          let col = '';
          let n = j + 1;
          while (n > 0) {
            let rem = (n - 1) % 26;
            col = String.fromCharCode(65 + rem) + col;
            n = Math.floor((n - 1) / 26);
          }
          coincidencias.push({
            fila: i + 1,
            columna: j + 1,
            referencia: `${col}${i + 1}`
          });
        }
      }
    }

    return coincidencias; // Devuelve array vacío si no hay resultados
  } catch (error) {
    throw new Error(`Error al buscar celdas: ${error.message}`);
  }
}


// ===========================
// EXPORTS
// ===========================

module.exports = {
  // Sheet y hojas
  crearSheet,
  crearHoja,
  buscarHoja,
  leerHoja,
  obtenerSheet,
  obtenerHoja,

  // Filas
  crearFila,
  leerFilas,
  buscarFila,
  buscarFilasPorCondiciones,
  leerFila,
  leerFilasPorIndices,
  borrarFila,

  // Columnas
  crearColumna,
  leerEncabezado,
  buscarColumna,
  obtenerColumna,
  borrarColumna,

  // Celdas
  llenarCelda,
  leerCelda,
  buscarCelda,
};