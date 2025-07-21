import fetch from 'node-fetch'; 

const PUENTE_URL = 'https://puente-sheets-jessica.onrender.com';

// Función genérica para enviar requests al puente
export async function enviarAccionAlPuente(accion, hoja, parametros) {
  const payload = { accion, hoja, parametros };

  try {
    const respuesta = await fetch(PUENTE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const data = await respuesta.json();

    if (data.status === 'ok') {
      console.log('✅ Acción ejecutada con éxito:', data.data);
      return data.data;
    } else {
      console.error('⚠️ Error del puente:', data);
      return null;
    }

  } catch (err) {
    console.error('❌ Error de conexión con el puente:', err.message);
    return null;
  }
}

// ✅ PRUEBA SIMPLE DEL CLIENTE
(async () => {
  const resultado = await enviarAccionAlPuente(
    'leerEncabezado',
    'Pendientes',
    {}
  );
  console.log('Resultado de prueba:', resultado);
})();