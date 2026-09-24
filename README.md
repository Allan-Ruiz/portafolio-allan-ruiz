# 📊 Portafolio Técnico | Allan Andre Ruiz Arenas
**Ingeniero Industrial (INACAP) | Data Operations & Business Intelligence | Procesos, Inventarios & Supply Chain**  
📍 Santiago (Peñalolén), Chile • 📧 [allanaruiz32@gmail.com](mailto:allanaruiz32@gmail.com) • 📱 +56 9 6857 9766

---

## 🎯 Perfil Profesional
Ingeniero Industrial especializado en **Análisis de Datos Operacionales (Data Operations)**, **Business Intelligence** y **Estandarización de Procesos e Inventarios**, complementado con formación práctica en infraestructura **AWS Cloud y Linux** (Generation Chile).

Experiencia comprobada diseñando pipelines ETL de limpieza y validación de datos (**Excel Avanzado, Power Query, BUSCARX, SQL, Power BI**), ejecutando migraciones de catálogos maestros (**800+ SKU en formato `.xls` hacia el Sistema POS Bsale**) y auditando discrepancias físico-contables contra sistemas **ERP corporativos**.

---

## 🏆 Proyectos Estrella (Data BI & Procesos)

### 1. Pipeline ETL, Limpieza de Datos y Migración de Catálogo Maestro — *Librerías BigPEN*
* **Rol:** Analista de Datos Operacionales (Proyecto Independiente) | *Nov. 2025 – Mar. 2026*
* **Stack:** `Excel Avanzado` · `Power Query (ETL)` · `BUSCARX` · `Sistema POS Bsale (.xls)` · `QA asistido por IA`
* **Problema:** Catálogo comercial disperso en planillas informales de **+1.200 productos** recibidas vía WhatsApp, con registros duplicados, celdas vacías, heterogeneidad de formatos y errores de asignación de SKU frente al producto físico real.
* **Metodología Técnica:**
  1. **Cotejo Masivo Asistido por IA:** Búsqueda automatizada de SKUs estándar de mercado usando Gemini Pro y Claude como herramientas de apoyo.
  2. **Validación Cruzada Determinística (`BUSCARX`):** Cruce relacional en Excel entre los resultados de ambos modelos para garantizar consistencia total entre fuentes.
  3. **Auditoría de Calidad y Gestión con Cliente:** Cruce contra la base original del cliente, detectando **5 productos con fallas críticas en el SKU**, levantando el hallazgo formalmente y obteniendo aprobación para su corrección.
  4. **Procesamiento ETL en Power Query:** Limpieza de nulos, supresión de espacios, tipado de datos y estructuración de la matriz final en el formato exacto **`.xls`** requerido por el **Sistema POS Bsale**.
* **Resultados Medibles:**
  * ⚡ **100 productos/hora** de velocidad de procesamiento logístico y ETL.
  * ✅ **800+ SKU finales** migrados a **Sistema POS Bsale** con **0% de pérdida de información** y reportes de conciliación con KPIs de rotación y cobertura.

---

### 2. Automatización de Liquidaciones Logísticas, Diseño BPM y Conciliación ERP — *Ambientes Limpios S.A.*
* **Rol:** Asistente de Control de Inventarios (Práctica Profesional) | *Ene. 2023 – Jun. 2023*
* **Stack:** `Excel Avanzado (Modelado Lógico)` · `BPM / Diagramación de Procesos` · `Auditoría ERP` · `Control de Mermas`
* **Metodología y Logros:**
  * **Automatización Financiera-Operativa:** Desarrollo de modelos con fórmulas lógicas condicionales en Excel para automatizar la liquidación mensual y pago a transportistas según tramos de vueltas diarias (**`\$25.000` / `\$55.000`**), eliminando el 100% del cálculo manual y asegurando la cuadratura presupuestaria.
  * **Estandarización de Procesos (BPM):** Levantamiento y diseño del diagrama de flujo operacional de la bodega a solicitud del Gerente de Procesos, siendo felicitado por jefatura al ser **validado como insumo técnico oficial para auditorías internas de control**.
  * **Auditoría Físico-Contable:** Ejecución de conteos cíclicos y conciliación sistemática de entradas y salidas físicas contra el **sistema ERP**, auditando discrepancias para asegurar la integridad del inventario.

---

### 3. Mapeo de Distribución Espacial (Matriz A1-P9) y Reducción de Tiempos — *Tricot*
* **Rol:** Reponedor de tienda part-time | *Dic. 2025 – Feb. 2026*
* **Metodología y Logros:**
  * Diseño e implementación por iniciativa propia de una **Matriz de Codificación Espacial por coordenadas (`A1-P9`)** según modelo, corte y talla entre sala y bodega.
  * **Impacto Estimado:** Reducción de los tiempos estimados de búsqueda y extracción de mercadería de **aprox. 10 minutos a menos de 1 minuto por artículo** en temporada de alta demanda.

---

### 4. Infraestructura Cloud Automatizada y Administración Linux — *Bootcamp AWS (Generation Chile)*
* **Rol:** Ingeniero en Formación Cloud / Operaciones TI | *2026 (320 horas cronológicas)*
* **Arquitectura Implementada:**
  * Despliegue de servidores web **Apache (HTTPD)** 100% automatizados en el arranque mediante scripts **Bash (User Data)** sobre instancias **AWS EC2 (Amazon Linux 2023, `t3.micro` a `t3.small`)**.
  * Seguridad perimetral con **Security Groups** de mínimo privilegio (HTTP puerto 80), gobernanza **AWS IAM** correlacionada con archivos críticos de Linux (`/etc/passwd`, `/etc/shadow`, `/etc/sudoers`), **Termination Protection** y escalado de volúmenes **EBS (8 GB a 10 GB)**.
