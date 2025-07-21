// puente.js
const path = require('path');
const fs = require('fs');
const dotenv = require('dotenv');

// Cargar dotenv desde baul.env
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

if (process.env.API_TOKEN === 'CLAVESHEETSJESSICA') {
  console.log('✅ API_TOKEN es correcto');
} else {
  console.warn('❌ API_TOKEN inesperado');
}

// Leer service-account.json
const serviceAccountPath = path.join(__dirname, 'service-account.json');
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


const spreadsheetId = process.env.SPREADSHEET_ID;

// Importar todo el módulo de Sheets.js en una constante global
const Sheets = require('./sheets');

const express = require('express');
const app = express();

// Middleware para parsear JSON
app.use(express.json());

// Endpoint simple para probar que está vivo
app.get('/', (req, res) => {
  res.send('✅ Servidor Puente activo y escuchando');
});

// POST /puente - Procesa comandos, solo buscarCelda implementado por ahora
app.post('/puente', async (req, res) => {
  console.log('📨 req.body completo:', req.body);
  const { comandos } = req.body;
  console.log('📥 comandos recibidos:', comandos);
  const resultados = [];

  for (const cmd of comandos) {
    console.log('➡️ procesando comando:', cmd);

    // ✅ Mapeo automático para indexFila → index
    if (cmd.args && typeof cmd.args === 'object') {
      if (cmd.args.indexFila !== undefined && cmd.args.index === undefined) {
        cmd.args.index = cmd.args.indexFila;
      }
    }

    const accion = cmd.accion;
    const funcionSheets = Sheets[accion];

    if (typeof funcionSheets === 'function') {
      try {
        console.log(`🚀 Ejecutando ${accion} con args:`, cmd.args);

        // Construir los argumentos en orden: siempre serviceAccount y spreadsheetId primeros
        const args = [serviceAccount, spreadsheetId];

        // Si se especifica hoja, agregarla como tercer argumento
        if (cmd.hoja) args.push(cmd.hoja);

        // Si hay argumentos adicionales, expandirlos
        if (cmd.args) {
          if (Array.isArray(cmd.args)) {
            args.push(...cmd.args);
          } else if (typeof cmd.args === 'object') {
            // Si es objeto, sus valores en orden
            args.push(...Object.values(cmd.args));
          } else {
            args.push(cmd.args);
          }
        }

        const salida = await funcionSheets(...args);

        console.log(`✅ ${accion} finalizó, salida:`, salida);

        resultados.push({
          accion,
          status: 'ok',
          salida
        });
      } catch (err) {
        resultados.push({
          accion,
          status: 'error',
          mensaje: err.message
        });
      }
    } else {
      resultados.push({
        accion,
        status: 'pendiente',
        mensaje: 'Acción no implementada en Sheets.js'
      });
    }
  }

  res.json({ resultados });
});

// Iniciar servidor en puerto 3000
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`🚀 Puente escuchando en http://localhost:${PORT}`);
});

// ✅ Importar el clienteJessica.js para ejecutar la prueba mínima automáticamente
import('./clienteJessica.js');
