// POST /api/ensayos-ocr
// Extrae datos de actas de ensayo de áridos (ESOCAN) en PDF digital

import pdfParse from 'pdf-parse';

export const config = { api: { bodyParser: { sizeLimit: '25mb' } } };

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Método no permitido' });

  try {
    const { pdfBase64 } = req.body;
    if (!pdfBase64) return res.status(400).json({ ok: false, error: 'Falta pdfBase64' });

    const buffer = Buffer.from(pdfBase64, 'base64');
    const data = await pdfParse(buffer);
    const text = data.text;

    const result = parsearActa(text);
    return res.status(200).json({ ok: true, data: result, texto: text });
  } catch (e) {
    console.error('ensayos-ocr error:', e.message);
    return res.status(500).json({ ok: false, error: e.message });
  }
}

function parsearActa(text) {
  const r = {};

  // Nº ACTA
  const mActa = text.match(/N[ºo°]\s*ACTA[\s\S]{0,30}?(\d{4}\/\d+)/i);
  if (mActa) r.num_acta = mActa[1];

  // Nº ALBARÁN / MUESTRA
  const mAlb = text.match(/MUESTRA[\s\S]{0,20}?(\.\d{4}\/\d+|\d{4}\/\d+)/i);
  if (mAlb) r.num_albaran = mAlb[1].replace(/^\./, '');

  // Fecha de toma
  const mToma = text.match(/Fecha de toma[:\s]+(\d{2}\/\d{2}\/\d{4})/i);
  if (mToma) r.fecha_toma = isoFecha(mToma[1]);

  // Fecha de acta / fin ensayos
  const mActaFecha = text.match(/FECHA DE ACTA[\s\S]{0,10}?(\d{2}\/\d{2}\/\d{4})/i)
    || text.match(/Fin de ensayos[:\s]+(\d{2}\/\d{2}\/\d{4})/i);
  if (mActaFecha) r.fecha_acta = isoFecha(mActaFecha[1]);

  // Tipo de material → fracción
  const mMat = text.match(/Tipo de material[:\s]+[ÁA]rido\s+([\d\/]+)/i);
  if (mMat) r.fraccion = mMat[1].trim();

  // Tipo de ensayo
  if (/granulometr/i.test(text)) r.tipo_ensayo = 'granulometria';
  else if (/equivalente de arena/i.test(text)) r.tipo_ensayo = 'eq_arena';
  else if (/contenido.*finos|finos.*tamiz/i.test(text)) r.tipo_ensayo = 'cont_finos';
  else if (/\u00edndice.*lajas|lajas/i.test(text)) r.tipo_ensayo = 'ind_lajas';
  else if (/caras.*fractura|fractura/i.test(text)) r.tipo_ensayo = 'caras_fractura';

  // Resultados granulometría — buscar tabla Tamiz/Pasa
  if (r.tipo_ensayo === 'granulometria') {
    const resultados = {};
    const tamices = ['8','6,3','6.3','4','2','1','0,5','0.5','0,25','0.25','0,125','0.125','0,063','0.063'];
    tamices.forEach(function(t) {
      const tNorm = t.replace(',','.');
      const re = new RegExp(t.replace('.','[.,]') + '\\s+(\\d+)', 'i');
      const m = text.match(re);
      if (m) resultados['gran_' + tNorm] = parseInt(m[1]);
    });
    r.resultados = resultados;
  }

  // Eq. Arena
  if (r.tipo_ensayo === 'eq_arena') {
    const mEq = text.match(/Equivalente.*?(\d+(?:[,.]\d+)?)\s*%/i);
    if (mEq) r.resultados = { eq_arena: parseFloat(mEq[1].replace(',','.')) };
  }

  // Contenido finos
  if (r.tipo_ensayo === 'cont_finos') {
    const mFi = text.match(/(?:finos|tamiz 0[,.]063).*?(\d+(?:[,.]\d+)?)\s*%/i);
    if (mFi) r.resultados = { cont_finos: parseFloat(mFi[1].replace(',','.')) };
  }

  r.estado = 'recogido';
  return r;
}

function isoFecha(ddmmyyyy) {
  const p = ddmmyyyy.split('/');
  if (p.length !== 3) return ddmmyyyy;
  return p[2] + '-' + p[1] + '-' + p[0];
}
