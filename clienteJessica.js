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