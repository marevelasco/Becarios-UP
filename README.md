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
| `auth.py` | Login de jefes de becarios (lee credenciales de Secrets) |
| `verificaciones.py` | Guarda/lee en Postgres las horas que un jefe confirma o corrige, y el historial de cada corrección |
| `excel_store.py` | Guarda/lee en Postgres el último Excel subido, para que sobreviva a un refresh |
| `config_store.py` | Guarda/lee en Postgres la fecha límite de cada semestre (para el aviso de urgencia) |
| `reportes.py` | Arma el reporte descargable (.xlsx) del cumplimiento de horas |
| `scripts/generar_password.py` | Genera el hash de contraseña de un jefe nuevo, para pegar en Secrets |
| `sample_data/ejemplo_becarios.xlsx` | Excel de muestra con datos ficticios para probar |
| `assets/logo_tree.png`, `assets/logo_full.png` | Logo oficial de la UP (recortado/optimizado) que usa la interfaz |

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

Si ya hay una base Postgres conectada (ver siguiente sección), el Excel que subas queda
guardado ahí — la próxima vez que alguien abra el dashboard (o si se refresca la página),
se carga solo, sin pedir subirlo de nuevo. Solo se guarda cuando subes un archivo real; el
botón **“Usar archivo de ejemplo”** es solo para pruebas y nunca sobreescribe el Excel real
guardado.

## Verificación de horas por jefes de becarios

Cada evento del Excel tiene un `ENCARGADO` (quien organizó esa actividad) y una lista de
`Becario(s)` que ayudaron, con sus horas. Ese `ENCARGADO` puede iniciar sesión en el
dashboard y, para cada evento suyo, **confirmar o corregir** las horas de cada becario que
le apoyó. La hora que el jefe guarda es la que a partir de ahí cuenta en el total del
becario (no la del Excel crudo) — así es como se "firma" que sí trabajó esas horas.

Cuando un jefe inicia sesión, el dashboard se divide en dos pestañas — **Dashboard**
(la vista de siempre) y **Verificar horas** (solo sus eventos) — para no mezclar todo en
un solo scroll largo. Ahí, cada evento es una fila plegable: si ya está 100% verificado
aparece cerrada con un resumen (ej. "3/3 verificadas"); si le falta algo, se abre sola.

Esto necesita dos cosas configuradas en `st.secrets` (nunca en el Excel ni en el repo):

1. **Credenciales de los jefes** — un usuario/contraseña por jefe. El campo `name` debe
   escribirse igual que aparece en la columna `ENCARGADO` del Excel (acentos/mayúsculas no
   importan). Genera el hash de la contraseña con:
   ```bash
   python3 scripts/generar_password.py
   ```
2. **Una base de datos Postgres** donde se guardan las verificaciones (independiente del
   Excel; la tabla se crea sola la primera vez que corre la app). Recomendamos
   [Neon](https://neon.tech) (gratis, sin tarjeta): crea un proyecto, copia el "connection
   string" de la pestaña *Connection Details* (usa el que empieza con `postgresql://`).

En `.streamlit/secrets.toml` (local, ya está en `.gitignore` — nunca se sube):

```toml
[auth]
cookie_name = "becarios_up_auth"
cookie_key = "una-frase-larga-y-aleatoria-que-tu-inventes"
cookie_expiry_days = 30

[jefes.mlopez]
name = "Mariana López"
password = "PEGA_AQUÍ_EL_HASH_QUE_TE_DIO_EL_SCRIPT"

[connections.postgresql]
url = "postgresql+psycopg2://usuario:password@host/basedatos?sslmode=require"
```

Si `st.secrets` no tiene jefes configurados, el login simplemente no aparece y el dashboard
funciona exactamente igual que antes (nadie ve ni puede tocar nada nuevo).

## Extras para dirección

- **Reporte descargable**: botón "📥 Descargar reporte" junto a la tabla de cumplimiento, con
  los mismos datos que se ven en pantalla en dos formatos — **Excel** (para filtrar/analizar)
  o **PDF** (una hoja horizontal, lista para imprimir o mandar por correo).
- **Historial de verificaciones**: cada vez que un jefe corrige una hora ya verificada, queda
  un registro permanente (nunca se borra ni se sobreescribe) de quién, cuándo y de cuánto a
  cuánto cambió. Si un evento×becario tiene más de una corrección, aparece un botón "🕓 Ver
  historial" en el panel de verificación del jefe.
- **Fecha límite del semestre**: opcional, se configura en el popover "📁 Archivo" con un
  selector de fecha. En cuanto se pone, los becarios que no han completado su meta muestran
  un aviso junto a su estatus — "🟠 Quedan N días" si faltan 14 días o menos, o "🔴 Venció
  hace N días" si ya pasó la fecha. Sin fecha configurada, el dashboard se ve igual que
  siempre.

## Subir a Streamlit Community Cloud (gratis)

1. **Crea un repositorio en GitHub** (público o privado) y sube el proyecto:
   ```bash
   git init
   git add app.py data_loader.py calculations.py auth.py verificaciones.py excel_store.py \
           config_store.py reportes.py requirements.txt README.md .gitignore sample_data/ \
           scripts/ assets/
   git commit -m "Dashboard de becarios UP"
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/becarios-up.git
   git push -u origin main
   ```
   ⚠️ **No subas el Excel real** con nombres de alumnos ni `.streamlit/secrets.toml`: el
   `.gitignore` ya los excluye a ambos; solo se versiona el ejemplo ficticio de
   `sample_data/`.
2. Entra a [share.streamlit.io](https://share.streamlit.io) e inicia sesión con tu
   cuenta de GitHub.
3. Pulsa **“Create app” → “Deploy a public app from GitHub”**, elige el repositorio,
   rama `main` y archivo principal `app.py`.
4. Pulsa **Deploy**. En un par de minutos tendrás una URL pública tipo
   `https://becarios-up.streamlit.app` para compartir con dirección.
5. Para habilitar la verificación de horas: en el panel de la app en Streamlit Cloud, ve a
   **Settings → Secrets** y pega ahí el mismo contenido que armaste en
   `.streamlit/secrets.toml` (ver sección anterior), con el connection string de tu propia
   base en Neon y las credenciales de tus jefes. Guarda — la app se reinicia sola con los
   Secrets nuevos.

Cada `git push` a `main` redespliega la app automáticamente. El Excel con datos reales
nunca vive en el repo (el `.gitignore` lo excluye): la encargada lo sube desde el
navegador cada vez que quiera actualizar los números. Si hay una base Postgres
conectada, ese Excel (y las verificaciones de los jefes) quedan guardados ahí — sobreviven
a que se refresque la página, se cierre el navegador o se reinicie/redespliegue la app. Sin
base configurada, el dashboard sigue funcionando igual que antes: hay que volver a subir el
Excel en cada sesión nueva.

## Checklist para quien vaya a tener la app en producción

Si alguien más va a operar la versión "real" (sus propias cuentas de GitHub/Streamlit/Neon),
esto es lo único que necesita hacer, ya con el código terminado:

- [ ] Recibir los archivos del proyecto (incluye `auth.py`, `verificaciones.py`,
      `excel_store.py`, `config_store.py`, `reportes.py`, `scripts/generar_password.py`, la
      carpeta `assets/` con el logo, y los cambios en `app.py`/`calculations.py`) y subirlos
      a su propio repositorio de
      GitHub.
- [ ] Crear su cuenta/proyecto gratis en [neon.tech](https://neon.tech) y copiar su
      connection string (no hace falta correr ningún SQL a mano, la app crea la tabla sola).
- [ ] Desplegar en Streamlit Community Cloud apuntando a su repo (pasos 2–4 de arriba).
- [ ] En **Settings → Secrets** de su app, pegar su connection string de Neon y las
      credenciales de sus jefes (usar `scripts/generar_password.py` para generar cada hash).
- [ ] Confirmar que el nombre (`name`) de cada jefe en Secrets está escrito igual que en la
      columna `ENCARGADO` de su Excel.

Nada de esto requiere volver a tocar código — es solo configuración desde sus propias
cuentas.

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
