# Bitacora 13 Upgrade

Esta carpeta deja preparada la Bitacora 13 para:

- importar trades desde capturas sin exponer claves en el navegador
- usar Firebase Hosting para abrirla desde cualquier dispositivo
- usar Firebase Functions como backend seguro de extraccion

## Despliegue

1. Instala Firebase CLI si no la tienes.
2. Entra en esta carpeta.
3. Instala dependencias de functions:

```powershell
cd functions
npm install
cd ..
```

4. Define la clave de Anthropic para Functions:

```powershell
firebase functions:secrets:set ANTHROPIC_API_KEY
```

Si prefieres variables de entorno clasicas en tu entorno de deploy:

```powershell
$env:ANTHROPIC_API_KEY="tu_api_key"
```

5. Despliega:

```powershell
firebase deploy --only hosting,functions
```

## Archivos relevantes

- `trades_import.html`: wizard de importacion profesional
- `trading_discrecional.html`: dashboard con importacion desde modal
- `bitacora_trade_import_api.js`: cliente compartido para llamar al backend
- `functions/index.js`: backend seguro para OCR/IA de capturas
