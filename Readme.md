# ContaFin - Agente Contable-Financiero para PYMEs

ContaFin es un asistente especializado en contabilidad, finanzas operativas y cumplimiento fiscal para pequeñas y medianas empresas. Esta aplicación utiliza un modelo de lenguaje para proporcionar análisis financieros detallados, generar informes y crear plantillas Excel adaptadas a las necesidades contables y financieras de las PYMEs.

## Características principales

- **Análisis financieros** personalizados según sector y tamaño de empresa
- **Informes periódicos** con recomendaciones estratégicas
- **Plantillas Excel** automatizadas para:
  - Flujo de caja
  - Nómina y cálculos laborales
  - Balance general
  - Estado de resultados
  - Análisis de punto de equilibrio
  - Ratios financieros
- **API REST** completa para integración con otros sistemas

## Requisitos

- Python 3.11+
- Dependencias listadas en `requirements.txt`

## Cómo desplegar en Render

### Método 1: Despliegue automatizado

1. Crea una cuenta en [Render](https://render.com/) si aún no la tienes
2. Haz clic en el botón "New" y selecciona "Blueprint"
3. Conecta tu repositorio de GitHub donde está alojado este código
4. Render detectará automáticamente el archivo `render.yaml` y configurará el servicio

### Método 2: Despliegue manual

1. En tu dashboard de Render, haz clic en "New" y selecciona "Web Service"
2. Conecta tu repositorio de GitHub
3. Configura el servicio con los siguientes parámetros:
   - **Name**: ContaFin (o el nombre que prefieras)
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
4. Añade las siguientes variables de entorno:
   - `OLLAMA_URL`: https://evaenespanol.loca.lt (o la URL de tu servicio Ollama)
   - `MODEL_NAME`: llama3:8b (o el modelo que desees utilizar)

## Variables de entorno

| Variable | Descripción | Valor por defecto |
|----------|-------------|-------------------|
| `PORT` | Puerto en el que se ejecutará la aplicación | 5000 |
| `OLLAMA_URL` | URL del servicio Ollama | https://evaenespanol.loca.lt |
| `MODEL_NAME` | Nombre del modelo de lenguaje a utilizar | llama3:8b |

## Endpoints de la API

| Endpoint | Método | Descripción |
|----------|--------|-------------|
| `/` | GET | Información general del servicio |
| `/chat` | POST | Interactuar con ContaFin mediante mensajes |
| `/report` | GET | Obtener el último informe financiero |
| `/reports` | GET | Listar todos los informes disponibles |
| `/generate-report` | POST | Solicitar un nuevo informe financiero |
| `/excel-template` | GET | Obtener plantilla Excel según tipo solicitado |
| `/excel-templates` | GET | Listar todas las plantillas Excel disponibles |
| `/reset` | POST | Reiniciar una sesión de conversación |
| `/health` | GET | Verificar estado del servicio |

## Ejemplos de uso

### Consultar a ContaFin

```bash
curl -X POST http://localhost:5000/chat \
  -H "Content-Type: application/json" \
  -d '{"message": "¿Cómo puedo calcular el punto de equilibrio para mi negocio?", "session_id": "usuario123"}'
```

### Obtener una plantilla Excel

```bash
curl -X GET http://localhost:5000/excel-template?type=flujo_caja -o flujo_caja.xlsx
```

### Generar un nuevo informe financiero

```bash
curl -X POST http://localhost:5000/generate-report
```

## Licencia

Este proyecto está bajo la licencia MIT. Ver el archivo LICENSE para más detalles.