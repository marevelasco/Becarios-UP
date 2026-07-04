# 🎓 Becarios UP — Dashboard de servicio becario

Dashboard para el área de servicio becario de la Universidad Panamericana: horas de
servicio de cada becario contra su meta (según % de beca), estatus de evaluaciones 360
y próximos eventos. Los datos se leen del Excel de eventos que se sube desde el propio
dashboard — nada está hardcodeado.

## Estructura

| Archivo | Qué hace |
|---|---|
| `app.py` | Interfaz Streamlit |
| `data_loader.py` | Lee y parsea los 3 bloques del Excel (eventos, tabla de becas, becarios activos) |
| `calculations.py` | Cálculo de horas, metas, estatus y advertencias |
| `sample_data/ejemplo_becarios.xlsx` | Excel de muestra con datos ficticios para probar |

## Correr localmente

```bash
# 1. Crear entorno e instalar dependencias (una sola vez)
python3 -m venv .venv
source .venv/bin/activate        # en Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 2. Arrancar el dashboard
streamlit run app.py
```

Se abre en `http://localhost:8501`. Sube el Excel de eventos (o pulsa
**“Usar archivo de ejemplo”**) y elige la hoja del semestre en el selector de arriba
a la derecha — el dashboard detecta automáticamente la hoja con el formato correcto.

## Subir a Streamlit Community Cloud (gratis)

1. **Crea un repositorio en GitHub** (público o privado) y sube el proyecto:
   ```bash
   git init
   git add app.py data_loader.py calculations.py requirements.txt README.md .gitignore sample_data/
   git commit -m "Dashboard de becarios UP"
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/becarios-up.git
   git push -u origin main
   ```
   ⚠️ **No subas el Excel real** con nombres de alumnos: el `.gitignore` ya excluye los
   `.xlsx` de la raíz; solo se versiona el ejemplo ficticio de `sample_data/`.
2. Entra a [share.streamlit.io](https://share.streamlit.io) e inicia sesión con tu
   cuenta de GitHub.
3. Pulsa **“Create app” → “Deploy a public app from GitHub”**, elige el repositorio,
   rama `main` y archivo principal `app.py`.
4. Pulsa **Deploy**. En un par de minutos tendrás una URL pública tipo
   `https://becarios-up.streamlit.app` para compartir con dirección.

Cada `git push` a `main` redespliega la app automáticamente. El Excel con datos reales
nunca vive en el repo: la encargada lo sube desde el navegador cada vez que quiera
actualizar los números (la sesión lo mantiene en memoria mientras la pestaña esté abierta).

## Formato esperado del Excel

Una hoja por semestre (ej. **"Propuesta de 2027-1"**) con tres bloques:

1. **Eventos** (encabezados en la fila 1): `EVENTO`, `Fecha`, `Hora`,
   `Horas Contabilizables`, `Lugar`, `Becario(s) *Nombre igual a la tabla*`, `ENCARGADO`.
   - Si `Horas Contabilizables` está vacía, se estima con el horario de `Hora`
     (ej. "4:00 pm a 7:00 pm" → 3 h); si tampoco se puede, cuenta 0 y se avisa.
   - `Becario(s)` acepta nombres separados por coma y "Todos"/"TEAM" para todo el equipo.
2. **Tabla de becas**: encabezados `% Beca` y `Horas requeridas`.
3. **Becarios activos**: `Nombre`, `% Beca`, `Rol`, `Semestre`, `Eval. 360`.

Los nombres en `Becario(s)` deben escribirse igual que en la tabla de becarios
(mayúsculas, acentos y espacios extra no importan). Los que no coincidan aparecen
en la sección de **advertencias** y sus horas no se cuentan.
