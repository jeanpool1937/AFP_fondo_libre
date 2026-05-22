/**
 * Endpoint serverless de Vercel - Sirve la configuración de API keys
 * desde variables de entorno seguras (nunca expuestas en el código fuente)
 */
export default function handler(req, res) {
    // Permitir solo peticiones GET
    if (req.method !== 'GET') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const groqApiKey = process.env.GROQ_API_KEY;

    if (!groqApiKey) {
        return res.status(404).json({ 
            error: 'API key no configurada',
            hint: 'Configura GROQ_API_KEY en las variables de entorno de Vercel'
        });
    }

    // Cabeceras de seguridad: no cachear respuestas con credenciales
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
    res.setHeader('Pragma', 'no-cache');

    return res.status(200).json({ 
        groq_api_key: groqApiKey 
    });
}
