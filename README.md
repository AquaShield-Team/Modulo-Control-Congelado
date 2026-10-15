# 🛡️ AquaShield - Control Congelado COMEX 2026

Plataforma inteligente de control documental, trazabilidad y gestión logística de embarques marítimos de salmón congelado para **AquaChile**.

---

## 🚀 Arquitectura del Sistema

El sistema opera bajo un modelo desacoplado **"Frontend Web Autónomo + Motor ETL Local Seguro"**:

* **Para el Equipo COMEX (Usuarios Normales)**:
  * Acceden directamente a la interfaz web (`index.html`) alojada en GitHub Pages o navegador.
  * **Cargar Planilla**: Permite subir y cruzar archivos Excel en memoria local de forma instantánea.
  * **Edición y Trazabilidad**: Modificación directa de analistas, navieras, contenedores, BL, fechas reales de borrador y valija DHL.
  * **Gestión de Cartera de Clientes**: Configuración de criterios de cierre, requerimiento de DHL y analista titular asignado.
  * **Descarga de Reportes**: Exportación oficial en formato Excel con columnas desagregadas y alertas visuales.

* **Para el Administrador (PC Local)**:
  * Ejecuta la extracción segura desde **Azure Synapse Analytics (Data Lake)** y SAP VF05N.
  * Ejecuta `Actualizar Datos.bat` para procesar la lógica de negocio y actualizar la línea base oficial en `data_control_congelado.js`.
  * Publica los datos consolidados en GitHub con `git push`.

---

## 📂 Estructura del Proyecto

```text
├── index.html                     # Interfaz de usuario (SPA responsive con tema claro/oscuro)
├── data_control_congelado.js      # Base de datos oficial consolidada
├── xlsx-js-style.min.js           # Librería para exportación estilizada a Excel
├── reglas_clientes.json           # Reglas de negocio y carteras especiales por cliente
├── Asignacion Clientes.xlsx       # Matriz de asignación analista-cliente
├── mapeo_equipos.xlsx             # Mapeo de analistas y equipos de trabajo
├── actualizar_datos.py            # Motor ETL de consolidación COMEX 2026
├── extraer_synapse.py             # Conector automatizado a Azure Synapse (SSO Azure AD)
├── consulta_synapse_congelado.sql # Query SQL optimizada con particionamiento de pedidos
├── Abrir Control Congelado.bat    # Acceso directo para abrir la plataforma en el navegador
└── Actualizar Datos.bat           # Script de 1-clic para sincronización total (Admin)
```

---

## 💻 Guía de Uso Rápido

### Para el Equipo (Usuarios Normales)
1. Abrir el archivo `index.html` en Google Chrome o Microsoft Edge (o acceder al enlace de GitHub Pages).
2. Para cruzar información con una planilla local, hacer clic en **"Subir Planilla / Actualizar"**.
3. Para ver o editar el detalle de un pedido, hacer clic en el número de pedido en la tabla.
4. Para exportar la planilla completa con formato y celdas pendientes resaltadas, hacer clic en **"Descargar Excel"**.

### Para el Administrador (Actualización de Base)
1. Ejecutar `Actualizar Datos.bat` en tu equipo.
2. Confirmar que la sincronización con Synapse y SAP finalice exitosamente.
3. Subir la actualización a GitHub:
   ```bash
   git add data_control_congelado.js reglas_clientes.json
   git commit -m "Actualización base de pedidos COMEX"
   git push
   ```

---

*AquaShield © 2026 - AquaChile S.A.*
