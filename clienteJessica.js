import fetch from 'node-fetch'; 

const PUENTE_URL = 'https://puente-sheets-jessica.onrender.com/puente';

// Función genérica para enviar requests al puente
export async function enviarAccionAlPuente(accion, hoja, parametros) {
  const payload = {
    comandos: [
      {
        accion,
        hoja,
        args: parametros
      }
    ]
  };

  try {
    const respuesta = await fetch(PUENTE_URL, {
      method: 'POST',
      headers: { 
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    const rawText = await respuesta.text();
    console.log('📜 Respuesta cruda del puente:', rawText);

    let data;
    try {
      data = JSON.parse(rawText);
    } catch (parseErr) {
      console.error('⚠️ No se pudo parsear JSON:', parseErr.message);
      return null;
    }

    if (data.resultados && Array.isArray(data.resultados)) {
      const primerResultado = data.resultados[0];
      if (primerResultado.status === 'ok') {
        console.log(`✅ Acción ${primerResultado.accion} ejecutada con éxito:`, primerResultado.salida);
        return primerResultado.salida;
      } else {
        console.error(`⚠️ Error en ${primerResultado.accion}:`, primerResultado.mensaje || 'Sin mensaje');
        return null;
      }
    } else if (data.status === 'ok') {
      console.log('✅ Acción ejecutada con éxito:', data.data);
      return data.data;
    } else {
      console.error('⚠️ Respuesta inesperada del puente:', data);
      return null;
    }

  } catch (err) {
    console.error('❌ Error de conexión con el puente:', err.message);
    return null;
  }
}

// ✅ PRUEBA INTEGRAL DE TODAS LAS FUNCIONES DISPONIBLES DEL PUENTE
async function pruebaIntegralFunciones() {
  console.log('🚀 Iniciando prueba integral de todas las funciones del puente');

  // BLOQUE 1 – SHEET Y HOJAS
  console.log('📌 BLOQUE 1 – Sheet y Hojas');
  const hojaExiste = await enviarAccionAlPuente('buscarHoja', 'Pendientes', { nombreHoja: 'Pendientes' });
  console.log('Resultado buscarHoja:', hojaExiste);

  const metadatosSheet = await enviarAccionAlPuente('obtenerSheet', 'Pendientes', {});
  console.log('Resultado obtenerSheet:', metadatosSheet ? 'OK' : 'No retornó');

  const metadatosHoja = await enviarAccionAlPuente('obtenerHoja', 'Pendientes', { nombreHoja: 'Pendientes' });
  console.log('Resultado obtenerHoja:', metadatosHoja ? 'OK' : 'No retornó');

  // BLOQUE 2 – FILAS
  console.log('📌 BLOQUE 2 – Filas');
  const filasLeidas = await enviarAccionAlPuente('leerFilas', 'Pendientes', {});
  console.log('Filas:', filasLeidas ? 'Leídas correctamente' : 'Error');

  // Ajuste correcto para leerFila por índice válido (2 = primera fila de datos reales)
  const filaPorIndice = await enviarAccionAlPuente('leerFila', 'Pendientes', { index: 2 });
  console.log('leerFila por índice 2 (primera fila de datos):', filaPorIndice);

  // Prueba leerFila por condición campo/valor
  const filaPorCondicion = await enviarAccionAlPuente('leerFila', 'Pendientes', {
    campo: 'ACTIVIDAD',
    valor: 'Sesión 1 Ciclo Lecciones'
  });
  console.log('leerFila por condición ACTIVIDAD=Sesión 1 Ciclo Lecciones:', filaPorCondicion);

  const filasPorCondicion = await enviarAccionAlPuente('buscarFilasPorCondiciones', 'Pendientes', {
    condiciones: [
      { campo: 'EJE', valor: 'Democracy Innovation Lab' }
    ]
  });
  console.log('Filas por condición EJE=Democracy Innovation Lab:', filasPorCondicion);

  // BLOQUE 3 – COLUMNAS
  console.log('📌 BLOQUE 3 – Columnas');
  const encabezados = await enviarAccionAlPuente('leerEncabezado', 'Pendientes', {});
  console.log('Encabezados:', encabezados);

  const buscarCol = await enviarAccionAlPuente('buscarColumna', 'Pendientes', { nombreColumna: 'EJE' });
  console.log('buscarColumna "EJE":', buscarCol);

  const valoresColumna = await enviarAccionAlPuente('obtenerColumna', 'Pendientes', { nombreColumna: 'EJE' });
  console.log('Valores de columna EJE:', valoresColumna);

  // BLOQUE 4 – CELDAS
  console.log('📌 BLOQUE 4 – Celdas');
  const celdaValor = await enviarAccionAlPuente('leerCelda', 'Pendientes', { celda: 'A2' });
  console.log('Valor celda A2:', celdaValor);

  const llenarCeldaTest = await enviarAccionAlPuente('llenarCelda', 'Pendientes', { celda: 'J2', valor: 'Prueba integral' });
  console.log('llenarCelda test:', llenarCeldaTest);

  const buscarCeldaTest = await enviarAccionAlPuente('buscarCelda', 'Pendientes', { valorBuscado: 'Prueba integral' });
  console.log('buscarCelda por valor "Prueba integral":', buscarCeldaTest);

  console.log('✅ Prueba integral completada');
}

// Ejecutar prueba integral
(async () => {
  await pruebaIntegralFunciones();
})();