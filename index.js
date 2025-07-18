const path = require('path');
const fs = require('fs');
const dotenv = require('dotenv');
const minimist = require('minimist');
const args = minimist(process.argv.slice(2));

(async function main() {

const envPath = path.join(__dirname, 'baul.env');

console.log('🛠 Verificando si baul.env existe en:', envPath);

if (!fs.existsSync(envPath)) {
  console.log('⛔ No se encontró el archivo baul.env');
  process.exit(1);
}

const result = dotenv.config({ path: envPath });

if (result.error) {
  console.log('⛔ Error al cargar baul.env:', result.error);
  process.exit(1);
}

console.log('✅ dotenv cargado desde baul.env');
console.log('🔑 process.env.PORT:', process.env.PORT);
console.log('🔑 process.env.API_TOKEN:', process.env.API_TOKEN);

const serviceAccountPath = path.join(__dirname, 'service-account.json');

if (process.env.API_TOKEN === 'CLAVESHEETSJESSICA') {
  console.log('✅ API_TOKEN es correcto');
} else {
  console.warn('❌ API_TOKEN inesperado');
}

// Verificar, leer y validar el archivo service-account.json
if (!fs.existsSync(serviceAccountPath)) {
  console.log('⛔ No se encontró el archivo service-account.json');
  process.exit(1);
}

const serviceAccountRaw = fs.readFileSync(serviceAccountPath, 'utf8');
let serviceAccount;

try {
  serviceAccount = JSON.parse(serviceAccountRaw);
} catch (err) {
  console.log('⛔ Error al parsear service-account.json:', err);
  process.exit(1);
}

if (!serviceAccount.client_email || !serviceAccount.private_key) {
  console.log('⛔ El archivo service-account.json no contiene client_email o private_key');
  process.exit(1);
}

console.log('✅ service-account.json cargado correctamente');
console.log('📧 client_email:', serviceAccount.client_email);

const { buscarHoja, leerHoja } = require('./sheets');

const spreadsheetId = process.env.SPREADSHEET_ID;

if (!args._[0]) {
  console.log('⛔ Debes especificar un argumento (nombre de hoja o acción):');
  console.log('   👉 Ejemplo: node index.js Test4Jessica');
  console.log('   👉 Ejemplo: node index.js crearFila');
  process.exit(1);
}

// --- Comando buscarCelda ---
if (args._[0] === 'buscarCelda') {
  const { buscarCelda } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.valor) {
    console.error('⛔ Debes proporcionar --valor para buscar en todas las celdas');
    process.exit(1);
  }

  try {
    const coincidencias = await buscarCelda(serviceAccount, spreadsheetId, hoja, args.valor);
    if (coincidencias.length === 0) {
      console.log(`⚠️ No se encontraron coincidencias para "${args.valor}" en la hoja "${hoja}".`);
    } else {
      console.log(`✅ Se encontraron ${coincidencias.length} coincidencia(s) para "${args.valor}":`);
      console.table(coincidencias);
    }
  } catch (errorBuscarCelda) {
    console.error('❌ Error al buscar la celda:', errorBuscarCelda.message);
  }

  return;
}

// --- Comando leerCelda ---
if (args._[0] === 'leerCelda') {
  const { leerCelda } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.celda) {
    console.error('⛔ Debes proporcionar --celda con la referencia exacta, ej: E5');
    process.exit(1);
  }

  try {
    const valor = await leerCelda(serviceAccount, spreadsheetId, hoja, args.celda);
    if (valor === null) {
      console.log(`⚠️ La celda ${args.celda} está vacía o no existe en la hoja "${hoja}".`);
    } else {
      console.log(`✅ Valor leído en ${args.celda}: "${valor}"`);
    }
  } catch (errorCelda) {
    console.error('❌ Error al leer la celda:', errorCelda.message);
  }

  return;
}

// --- Comando llenarCelda ---
if (args._[0] === 'llenarCelda') {
  const { llenarCelda } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.celda || args.valor === undefined) {
    console.error('⛔ Debes proporcionar --celda (ej: E5) y --valor');
    process.exit(1);
  }

  try {
    await llenarCelda(serviceAccount, spreadsheetId, hoja, args.celda, args.valor);
    console.log(`✅ Celda ${args.celda} actualizada correctamente con el valor: "${args.valor}"`);
  } catch (errorLlenarCelda) {
    console.error('❌ Error al llenar la celda:', errorLlenarCelda.message);
  }

  return;
}

if (args._[0] === 'crearFila') {
  const { crearFila } = require('./sheets');

  const datos = {
    FECHA: args.FECHA || '',
    EJE: args.EJE || '',
    PROYECTO: args.PROYECTO || '',
    SUBPROYECTO: args.SUBPROYECTO || '',
    TEMA: args.TEMA || '',
    DETALLE: args.DETALLE || '',
    ESTADO: args.ESTADO || '',
    PRIORIDAD: args.PRIORIDAD || '',
    ASIGNADO_A: args.ASIGNADO_A || '',
    OBSERVACIONES: args.OBSERVACIONES || ''
  };

  try {
    await crearFila(serviceAccount, spreadsheetId, args.hoja || 'Test4Jessica', datos);
    console.log('✅ Fila insertada correctamente en la hoja');
  } catch (errorCreacion) {
    console.error('❌ Error al crear la fila:', errorCreacion.message);
  }

  return;
}

if (args._[0] === 'crearFilas') {
  const { crearFila } = require('./sheets');

  if (!args.filas) {
    console.error('⛔ Debes proporcionar el argumento --filas con un JSON válido');
    process.exit(1);
  }

  let registros;
  try {
    registros = JSON.parse(args.filas);
    if (!Array.isArray(registros)) throw new Error('El contenido de --filas debe ser un arreglo');
  } catch (err) {
    console.error('❌ Error al parsear el argumento --filas:', err.message);
    process.exit(1);
  }

  for (const fila of registros) {
    try {
      await crearFila(serviceAccount, spreadsheetId, args.hoja || 'Test4Jessica', fila);
      console.log(`✅ Fila insertada: ${fila.TEMA || '[sin tema]'}`);
    } catch (errorFila) {
      console.error(`❌ Error al insertar fila ${fila.TEMA || '[sin tema]'}:`, errorFila.message);
    }
  }

  return;
}

// --- Comando leerFila ---
if (args._[0] === 'leerFila') {
  const { leerFila } = require('./sheets');

  const hoja = args.hoja || 'Test4Jessica';

  const opciones = {
    index: args.index,
    campo: args.campo,
    valor: args.valor
  };

  try {
    const fila = await leerFila(serviceAccount, spreadsheetId, hoja, opciones);
    console.log('📍 Fila encontrada:');
    console.table([fila]);
  } catch (errorLectura) {
    console.error('❌ Error al leer la fila:', errorLectura.message);
  }

  return;
}

// --- Comando leerFilasCondiciones ---
if (args._[0] === 'leerFilasCondiciones') {
  const { buscarFilasPorCondiciones } = require('./sheets');

  const hoja = args.hoja || 'Test4Jessica';

  // Recibir condiciones desde argumentos CLI
  // Ejemplo: --condiciones '[{"campo":"ESTADO","valor":"En proceso"},{"campo":"ASIGNADO_A","valor":"Demóstenes"}]'
  if (!args.condiciones) {
    console.error('⛔ Debes proporcionar --condiciones con un JSON válido de condiciones');
    process.exit(1);
  }

  let condiciones;
  try {
    condiciones = JSON.parse(args.condiciones);
    if (!Array.isArray(condiciones)) throw new Error('El contenido de --condiciones debe ser un arreglo');
  } catch (err) {
    console.error('❌ Error al parsear el argumento --condiciones:', err.message);
    process.exit(1);
  }

  try {
    const resultados = await buscarFilasPorCondiciones(serviceAccount, spreadsheetId, hoja, condiciones);
    if (resultados.length === 0) {
      console.log('⚠️ No se encontraron filas que coincidan con las condiciones dadas.');
    } else {
      console.log(`✅ Se encontraron ${resultados.length} fila(s) que cumplen las condiciones:`);
      console.table(resultados);
    }
  } catch (errorBusqueda) {
    console.error('❌ Error al buscar filas por condiciones:', errorBusqueda.message);
  }

  return;
}

// --- Comando leerFilasIndices ---
if (args._[0] === 'leerFilasIndices') {
  const { leerFilasPorIndices } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.indices) {
    console.error('⛔ Debes proporcionar --indices como JSON de números, ej: --indices "[2,5,9]"');
    process.exit(1);
  }

  let indices;
  try {
    indices = JSON.parse(args.indices);
    if (!Array.isArray(indices)) throw new Error('El contenido de --indices debe ser un arreglo');
  } catch (err) {
    console.error('❌ Error al parsear --indices:', err.message);
    process.exit(1);
  }

  try {
    const resultados = await leerFilasPorIndices(serviceAccount, spreadsheetId, hoja, indices);
    console.log(`✅ Resultados para índices ${indices.join(', ')}:`);
    console.table(resultados);
  } catch (errorLectura) {
    console.error('❌ Error al leer filas por índices:', errorLectura.message);
  }

  return;
}

// --- Comando borrarFila ---
if (args._[0] === 'borrarFila') {
  const { borrarFila } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.fila) {
    console.error('⛔ Debes proporcionar --fila con el número de fila a borrar');
    process.exit(1);
  }

  const numeroFila = parseInt(args.fila, 10);
  if (isNaN(numeroFila) || numeroFila <= 0) {
    console.error('❌ El argumento --fila debe ser un número válido mayor a 0');
    process.exit(1);
  }

  try {
    await borrarFila(serviceAccount, spreadsheetId, hoja, numeroFila);
    console.log(`✅ Fila ${numeroFila} borrada correctamente de la hoja "${hoja}"`);
  } catch (errorBorrar) {
    console.error('❌ Error al borrar la fila:', errorBorrar.message);
  }

  return;
}

// --- Comando borrarFilaCondicion ---
if (args._[0] === 'borrarFilaCondicion') {
  const { buscarFilasPorCondiciones, leerFilas, borrarFila } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.condiciones) {
    console.error('⛔ Debes proporcionar --condiciones con un JSON válido de condiciones');
    process.exit(1);
  }

  let condiciones;
  try {
    condiciones = JSON.parse(args.condiciones);
    if (!Array.isArray(condiciones)) throw new Error('El contenido de --condiciones debe ser un arreglo');
  } catch (err) {
    console.error('❌ Error al parsear --condiciones:', err.message);
    process.exit(1);
  }

  try {
    // Buscar filas que cumplan
    const resultados = await buscarFilasPorCondiciones(serviceAccount, spreadsheetId, hoja, condiciones);
    if (resultados.length === 0) {
      console.log('⚠️ No se encontraron filas que coincidan con las condiciones dadas.');
      return;
    }

    // Obtener todas las filas originales
    const filas = await leerFilas(serviceAccount, spreadsheetId, hoja);
    const encabezados = filas[0];
    const filasSinEncabezado = filas.slice(1);

    // Tomar la primera coincidencia encontrada por condiciones
    const filaObjetivo = resultados[0];

    // Determinar el índice relativo en las filas de datos
    const idxEnDatos = filasSinEncabezado.findIndex(fila =>
      encabezados.every((col, colIdx) => (filaObjetivo[col] || '') === (fila[colIdx] || ''))
    );

    if (idxEnDatos === -1) {
      console.log('⚠️ La fila encontrada no se pudo mapear a un índice real.');
      return;
    }

    // Ahora fila 1 = primera fila de datos
    const idxUsuario = idxEnDatos + 1;

    // Pasar índice usuario directamente, borrarFila ajusta internamente
    await borrarFila(serviceAccount, spreadsheetId, hoja, idxUsuario);
    console.log(`✅ Fila borrada correctamente (fila usuario: ${idxUsuario})`);

  } catch (errorCond) {
    console.error('❌ Error al borrar fila por condición:', errorCond.message);
  }

  return;
}

// --- Comando borrarFilasIndices ---
if (args._[0] === 'borrarFilasIndices') {
  const { borrarFila } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.indices) {
    console.error('⛔ Debes proporcionar --indices como JSON de números, ej: --indices "[2,5,9]"');
    process.exit(1);
  }

  let indices;
  try {
    indices = JSON.parse(args.indices);
    if (!Array.isArray(indices)) throw new Error('El contenido de --indices debe ser un arreglo');
  } catch (err) {
    console.error('❌ Error al parsear --indices:', err.message);
    process.exit(1);
  }

  // Ordenar de mayor a menor para evitar desfase de borrado
  indices.sort((a, b) => b - a);

  try {
    for (const idx of indices) {
      const numeroFila = parseInt(idx, 10);
      if (isNaN(numeroFila) || numeroFila <= 0) {
        console.warn(`⚠️ Índice inválido ${idx}, se omite.`);
        continue;
      }
      await borrarFila(serviceAccount, spreadsheetId, hoja, numeroFila);
      console.log(`✅ Fila ${numeroFila} borrada correctamente de la hoja "${hoja}"`);
    }
    console.log('✅ Borrado múltiple completado.');
  } catch (errorMulti) {
    console.error('❌ Error al borrar filas múltiples:', errorMulti.message);
  }

  return;
}

// --- Comando borrarFilasCondiciones ---
if (args._[0] === 'borrarFilasCondiciones') {
  const { buscarFilasPorCondiciones, leerFilas, borrarFila } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.condiciones) {
    console.error('⛔ Debes proporcionar --condiciones con un JSON válido de condiciones');
    process.exit(1);
  }

  let condiciones;
  try {
    condiciones = JSON.parse(args.condiciones);
    if (!Array.isArray(condiciones)) throw new Error('El contenido de --condiciones debe ser un arreglo');
  } catch (err) {
    console.error('❌ Error al parsear --condiciones:', err.message);
    process.exit(1);
  }

  try {
    // Buscar todas las filas que cumplan
    const resultados = await buscarFilasPorCondiciones(serviceAccount, spreadsheetId, hoja, condiciones);
    if (resultados.length === 0) {
      console.log('⚠️ No se encontraron filas que coincidan con las condiciones dadas.');
      return;
    }

    // Obtener todas las filas para mapear índices
    const filas = await leerFilas(serviceAccount, spreadsheetId, hoja);
    const encabezados = filas[0];
    const filasSinEncabezado = filas.slice(1);

    // Mapear índices de usuario para cada coincidencia
    const indicesUsuario = resultados.map(filaObjetivo => {
      const idxEnDatos = filasSinEncabezado.findIndex(fila =>
        encabezados.every((col, colIdx) => (filaObjetivo[col] || '') === (fila[colIdx] || ''))
      );
      return idxEnDatos !== -1 ? idxEnDatos + 1 : null; // +1 porque fila 1 = primera fila de datos
    }).filter(idx => idx !== null);

    if (indicesUsuario.length === 0) {
      console.log('⚠️ No se pudieron mapear índices para las coincidencias.');
      return;
    }

    // Ordenar de mayor a menor para evitar desfases
    indicesUsuario.sort((a, b) => b - a);

    // Borrar todas las filas encontradas
    for (const idxUsuario of indicesUsuario) {
      await borrarFila(serviceAccount, spreadsheetId, hoja, idxUsuario);
      console.log(`✅ Fila borrada (fila usuario: ${idxUsuario})`);
    }

    console.log(`✅ Se borraron ${indicesUsuario.length} fila(s) que cumplían las condiciones.`);

  } catch (errorCondMulti) {
    console.error('❌ Error al borrar filas por condiciones:', errorCondMulti.message);
  }

  return;
}

// --- Comando borrarFilasMixto ---
if (args._[0] === 'borrarFilasMixto') {
  const { buscarFilasPorCondiciones, leerFilas, borrarFila } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.indices && !args.condiciones) {
    console.error('⛔ Debes proporcionar al menos --indices o --condiciones');
    process.exit(1);
  }

  // Procesar indices
  let indicesDirectos = [];
  if (args.indices) {
    try {
      indicesDirectos = JSON.parse(args.indices);
      if (!Array.isArray(indicesDirectos)) throw new Error('El contenido de --indices debe ser un arreglo');
    } catch (err) {
      console.error('❌ Error al parsear --indices:', err.message);
      process.exit(1);
    }
  }

  // Procesar condiciones
  let condiciones = [];
  if (args.condiciones) {
    try {
      condiciones = JSON.parse(args.condiciones);
      if (!Array.isArray(condiciones)) throw new Error('El contenido de --condiciones debe ser un arreglo');
    } catch (err) {
      console.error('❌ Error al parsear --condiciones:', err.message);
      process.exit(1);
    }
  }

  try {
    let indicesPorCondiciones = [];
    if (condiciones.length > 0) {
      const resultados = await buscarFilasPorCondiciones(serviceAccount, spreadsheetId, hoja, condiciones);
      if (resultados.length > 0) {
        const filas = await leerFilas(serviceAccount, spreadsheetId, hoja);
        const encabezados = filas[0];
        const filasSinEncabezado = filas.slice(1);

        indicesPorCondiciones = resultados.map(filaObjetivo => {
          const idxEnDatos = filasSinEncabezado.findIndex(fila =>
            encabezados.every((col, colIdx) => (filaObjetivo[col] || '') === (fila[colIdx] || ''))
          );
          return idxEnDatos !== -1 ? idxEnDatos + 1 : null; 
        }).filter(idx => idx !== null);
      }
    }

    // Unir todos los índices
    let todosLosIndices = [...indicesDirectos, ...indicesPorCondiciones];
    if (todosLosIndices.length === 0) {
      console.log('⚠️ No se encontraron filas que coincidan ni índices válidos.');
      return;
    }

    // Eliminar duplicados y ordenar descendente
    todosLosIndices = [...new Set(todosLosIndices)].sort((a, b) => b - a);

    for (const idxUsuario of todosLosIndices) {
      await borrarFila(serviceAccount, spreadsheetId, hoja, idxUsuario);
      console.log(`✅ Fila borrada (fila usuario: ${idxUsuario})`);
    }

    console.log(`✅ Se borraron ${todosLosIndices.length} fila(s) por combinación de índices y condiciones.`);

  } catch (errorMixto) {
    console.error('❌ Error al borrar filas mixto:', errorMixto.message);
  }

  return;
}

// --- Comando leerColumna ---
if (args._[0] === 'leerColumna') {
  const { obtenerColumna } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.columna) {
    console.error('⛔ Debes proporcionar --columna con el nombre exacto del encabezado');
    process.exit(1);
  }

  try {
    const valores = await obtenerColumna(serviceAccount, spreadsheetId, hoja, args.columna);
    if (!valores.length) {
      console.log(`⚠️ La columna "${args.columna}" no tiene datos (o no existe).`);
    } else {
      console.log(`✅ Valores de la columna "${args.columna}":`);
      console.table(valores);
    }
  } catch (errorCol) {
    console.error('❌ Error al leer la columna:', errorCol.message);
  }

  return;
}

// --- Comando leerEncabezado ---
if (args._[0] === 'leerEncabezado') {
  const { leerEncabezado } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  try {
    const encabezados = await leerEncabezado(serviceAccount, spreadsheetId, hoja);
    if (!encabezados || encabezados.length === 0) {
      console.log('⚠️ No se encontraron encabezados en la hoja.');
    } else {
      console.log('✅ Encabezados detectados:');
      console.table(encabezados);
    }
  } catch (errorCols) {
    console.error('❌ Error al leer los encabezados:', errorCols.message);
  }

  return;
}

// --- Comando crearColumnaPosicion ---
if (args._[0] === 'crearColumnaPosicion') {
  const { leerFilas } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.tituloColumna || !args.posicion) {
    console.error('⛔ Debes proporcionar --tituloColumna y --posicion');
    process.exit(1);
  }

  const tituloColumna = args.tituloColumna;
  const posicion = parseInt(args.posicion, 10);

  if (isNaN(posicion) || posicion <= 0) {
    console.error('❌ La posición debe ser un número válido mayor a 0');
    process.exit(1);
  }

  try {
    const filas = await leerFilas(serviceAccount, spreadsheetId, hoja);
    if (!filas.length) throw new Error('La hoja no tiene filas para insertar columna');

    // Insertar en la posición indicada (1 = primera columna)
    const idxInsertar = posicion - 1;

    for (let i = 0; i < filas.length; i++) {
      if (i === 0) {
        // Encabezados: insertar título
        filas[i].splice(idxInsertar, 0, tituloColumna);
      } else {
        // Filas de datos: insertar celda vacía
        filas[i].splice(idxInsertar, 0, '');
      }
    }

    // Sobrescribir todo el rango con la nueva estructura
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${hoja}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: filas }
    });

    console.log(`✅ Columna "${tituloColumna}" insertada correctamente en posición ${posicion}`);

  } catch (errorPos) {
    console.error('❌ Error al crear columna por posición:', errorPos.message);
  }

  return;
}

// --- Comando crearColumnaReferencia ---
if (args._[0] === 'crearColumnaReferencia') {
  const { leerFilas } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.tituloColumna || !args.referencia || !args.posicionRef) {
    console.error('⛔ Debes proporcionar --tituloColumna, --referencia y --posicionRef ("antes" o "despues")');
    process.exit(1);
  }

  const tituloColumna = args.tituloColumna;
  const referencia = args.referencia;
  const posicionRef = args.posicionRef.toLowerCase();

  try {
    const filas = await leerFilas(serviceAccount, spreadsheetId, hoja);
    if (!filas.length) throw new Error('La hoja no tiene filas para insertar columna');

    const encabezados = filas[0];
    const idxRef = encabezados.indexOf(referencia);

    if (idxRef === -1) {
      console.error(`❌ La columna de referencia "${referencia}" no existe en la hoja.`);
      process.exit(1);
    }

    // Calcular índice de inserción según la posición solicitada
    let idxInsertar = posicionRef === 'antes' ? idxRef : idxRef + 1;

    for (let i = 0; i < filas.length; i++) {
      if (i === 0) {
        filas[i].splice(idxInsertar, 0, tituloColumna);
      } else {
        filas[i].splice(idxInsertar, 0, '');
      }
    }

    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: serviceAccount,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: authClient });

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${hoja}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: filas }
    });

    console.log(`✅ Columna "${tituloColumna}" insertada ${posicionRef} de "${referencia}"`);

  } catch (errorRef) {
    console.error('❌ Error al crear columna por referencia:', errorRef.message);
  }

  return;
}

// --- Comando buscarColumna ---
if (args._[0] === 'buscarColumna') {
  const { buscarColumna } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.columna) {
    console.error('⛔ Debes proporcionar --columna con el nombre exacto del encabezado');
    process.exit(1);
  }

  try {
    const indice = await buscarColumna(serviceAccount, spreadsheetId, hoja, args.columna);
    if (indice === -1) {
      console.log(`❌ La columna "${args.columna}" no se encontró en la hoja "${hoja}".`);
    } else {
      console.log(`✅ La columna "${args.columna}" existe en índice ${indice} (basado en 0 para encabezados).`);
    }
  } catch (errorBuscarCol) {
    console.error('❌ Error al buscar la columna:', errorBuscarCol.message);
  }

  return;
}

// --- Comando borrarColumna ---
if (args._[0] === 'borrarColumna') {
  const { borrarColumna } = require('./sheets');
  const hoja = args.hoja || 'Test4Jessica';

  if (!args.columna) {
    console.error('⛔ Debes proporcionar --columna con el nombre exacto a borrar');
    process.exit(1);
  }

  try {
    await borrarColumna(serviceAccount, spreadsheetId, hoja, args.columna);
    console.log(`✅ Columna "${args.columna}" borrada correctamente de la hoja "${hoja}"`);
  } catch (errorBorrarColumna) {
    console.error('❌ Error al borrar la columna:', errorBorrarColumna.message);
  }

  return;
}

const hojaABuscar = args._[0];

console.log('🧾 ID de hoja tomado desde baul.env:', spreadsheetId);
console.log(`🔍 Buscando hoja "${hojaABuscar}" en el documento...`);

console.log('🔍 Preparando autenticación para obtener listado de hojas...');

try {
  const existe = await buscarHoja(serviceAccount, spreadsheetId, hojaABuscar);
  if (existe) {
    console.log(`✅ La hoja "${hojaABuscar}" existe`);
  } else {
    console.log(`❌ La hoja "${hojaABuscar}" NO existe`);
  }

  // Si la hoja existe, se procede a leerla
  if (existe) {
    console.log(`📖 Leyendo contenido de la hoja "${hojaABuscar}"...`);
    try {
      const datos = await leerHoja(serviceAccount, spreadsheetId, hojaABuscar);
      console.log('✅ Contenido leído (crudo):');
      console.dir(datos, { depth: null });
      console.table(datos);
    } catch (errorLectura) {
      console.error(`❌ Error al leer la hoja "${hojaABuscar}":`, errorLectura.message);
    }
  }
} catch (err) {
  console.error(`⛔ Error al buscar la hoja "${hojaABuscar}":`, err);
}
})().catch(err => console.error('⛔ Error inesperado en ejecución:', err));