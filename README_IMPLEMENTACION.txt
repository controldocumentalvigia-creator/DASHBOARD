CENTRO DE CONTROL GERENCIAL VSE
VERSIÓN FINAL: EJECUTIVA + BASE PERSISTENTE COMPARTIDA

QUÉ CAMBIÓ
1. El dashboard abre de inmediato con la última base vigente.
2. Ya NO obliga a cargar un Excel para poder visualizar los datos.
3. "Actualizar base vigente" queda como una acción opcional en el panel lateral.
4. Cuando un jefe publica una nueva base, la app la guarda en GitHub y pasa a ser
   la nueva versión visible para todos.
5. Los dashboards OPERATIVO y FINANCIERO usan el mismo diseño Dark Navy ejecutivo.
6. Se corrigió el contraste de textos, tarjetas, filtros, tablas y gráficas.
7. No se implementó EBITDA porque la base actual no contiene los gastos necesarios.
8. No se cambiaron las fórmulas base de Producción, Costo, Margen, Rentabilidad,
   Recaudo, Pago a Terceros ni Caja Operativa.

ORDEN DASHBOARD FINANCIERO
1. Resultado económico
2. Recaudo y obligaciones
3. Posición de caja operativa
4. Evolución financiera mensual
5. Análisis por cliente y terceros
6. Proyección estadística próximo mes
7. Descarga ejecutiva

ORDEN DASHBOARD OPERATIVO
1. Resumen ejecutivo
2. Seguimiento de estatus operativo
3. Operación y recursos
4. Seguimiento de recaudo asociado a la producción
5. Matriz Cliente x Periodo
6. Análisis económico por placa y conductor
7. Servicios por placa y cliente

CONFIGURACIÓN ÚNICA PARA QUE LA CARGA SEA COMPARTIDA
La app necesita permiso para reemplazar el Excel maestro en GitHub.

En Streamlit Cloud:
1. Abre la aplicación.
2. Settings > Secrets.
3. Agrega:

GITHUB_TOKEN = "TU_TOKEN"
GITHUB_REPO = "USUARIO_O_ORGANIZACION/SEGUIMIENTO_OP_VSE"
GITHUB_BRANCH = "main"
GITHUB_MASTER_PATH = "TRAYECTOS TODOS__2026.xlsx"

El token debe tener permiso de lectura/escritura sobre CONTENTS del repositorio.
NO subir el token a GitHub.

COMPORTAMIENTO
- Sin cargar nada: abre la última base vigente y muestra el dashboard.
- Un jefe entra a "Actualizar base vigente", selecciona el .xlsx y la app valida:
  hoja TODOS, columnas obligatorias y fechas CARGA.
- Sólo después de validar se habilita "Publicar como base vigente".
- Al publicar: reemplaza el Excel maestro en GitHub y recalcula el dashboard.
- Otros usuarios que ingresen verán la última versión compartida.

RESPALDO
Si todavía no se configuran los Secrets, la app puede abrir el Excel maestro que
ya esté incluido en el repositorio. Sin Secrets, la actualización compartida desde
el dashboard permanece deshabilitada para no crear una falsa sensación de persistencia.

ARCHIVOS A SUBIR/REEMPLAZAR
- app.py
- requirements.txt
- .streamlit/config.toml

No es obligatorio subir secrets.toml.example; es sólo una guía.
Mantener en el repositorio el Excel maestro vigente:
TRAYECTOS TODOS__2026.xlsx
