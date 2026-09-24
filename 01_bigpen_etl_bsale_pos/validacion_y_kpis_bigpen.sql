-- ============================================================================
-- PROYECTO 1: PIPELINE DE VALIDACIÓN, AUDITORÍA DE SKUs Y KPIs DE INVENTARIO
-- Caso de Estudio: Migración Catálogo Maestro Librerías BigPEN -> Sistema POS Bsale
-- Autor: Allan Andre Ruiz Arenas | Ingeniero Industrial (INACAP)
-- ============================================================================

-- 1. CREACIÓN DE TABLAS DE STAGING (Simulación del flujo Excel / Power Query en SQL)
CREATE TABLE IF NOT EXISTS staging_planilla_cliente (
    id_registro INT PRIMARY KEY,
    sku_cliente VARCHAR(50),
    nombre_producto_raw VARCHAR(255),
    categoria VARCHAR(100),
    stock_fisico INT,
    costo_unitario DECIMAL(10, 2),
    precio_venta DECIMAL(10, 2)
);

CREATE TABLE IF NOT EXISTS staging_verificacion_ia (
    nombre_normalizado VARCHAR(255) PRIMARY KEY,
    sku_modelo_gemini VARCHAR(50),
    sku_modelo_claude VARCHAR(50),
    fuente_mercado VARCHAR(100)
);

-- 2. VALIDACIÓN CRUZADA DETERMINÍSTICA (Equivalente en SQL al cruce con BUSCARX)
-- Detecta dónde coinciden ambos modelos de apoyo y aísla las 5 anomalías críticas
-- donde el SKU histórico del cliente NO corresponde al artículo físico real.
WITH validacion_cruzada AS (
    SELECT 
        c.id_registro,
        TRIM(UPPER(c.nombre_producto_raw)) AS nombre_limpio,
        c.sku_cliente,
        v.sku_modelo_gemini,
        v.sku_modelo_claude,
        c.stock_fisico,
        c.costo_unitario,
        c.precio_venta,
        CASE 
            WHEN v.sku_modelo_gemini = v.sku_modelo_claude 
                 AND c.sku_cliente = v.sku_modelo_gemini THEN 'OK - VALIDADO'
            WHEN v.sku_modelo_gemini = v.sku_modelo_claude 
                 AND c.sku_cliente <> v.sku_modelo_gemini THEN 'ALERTA CRÍTICA: SKU CLIENTE ERRÓNEO (Escalar por WhatsApp)'
            ELSE 'REVISIÓN MANUAL REQUERIDA'
        END AS estado_auditoria_qa,
        COALESCE(v.sku_modelo_gemini, c.sku_cliente) AS sku_final_aprobado
    FROM staging_planilla_cliente c
    LEFT JOIN staging_verificacion_ia v
        ON TRIM(UPPER(c.nombre_producto_raw)) = TRIM(UPPER(v.nombre_normalizado))
    WHERE c.nombre_producto_raw IS NOT NULL
)
SELECT 
    id_registro,
    nombre_limpio,
    sku_cliente AS sku_original_erroneo,
    sku_final_aprobado AS sku_corregido_bsale,
    estado_auditoria_qa
FROM validacion_cruzada
WHERE estado_auditoria_qa LIKE 'ALERTA CRÍTICA%';

-- 3. VISTA FINAL ESTRUCTURADA PARA EXPORTACIÓN .XLS HACIA SISTEMA POS BSALE
-- Elimina duplicados, suprime espacios en blanco y normaliza tipos de datos
CREATE VIEW vw_matriz_importacion_bsale_xls AS
SELECT DISTINCT
    COALESCE(v.sku_modelo_gemini, TRIM(c.sku_cliente)) AS codigo_sku_bsale,
    TRIM(c.nombre_producto_raw) AS nombre_producto,
    UPPER(TRIM(COALESCE(c.categoria, 'LIBRERIA GENERAL'))) AS tipo_producto,
    GREATEST(COALESCE(c.stock_fisico, 0), 0) AS stock_inicial_bodega,
    ROUND(c.costo_unitario, 0) AS costo_neto_clp,
    ROUND(c.precio_venta, 0) AS precio_venta_bruto_clp
FROM staging_planilla_cliente c
LEFT JOIN staging_verificacion_ia v
    ON TRIM(UPPER(c.nombre_producto_raw)) = TRIM(UPPER(v.nombre_normalizado))
WHERE TRIM(c.nombre_producto_raw) <> '';

-- 4. REPORTE DE CONCILIACIÓN Y KPIs PARA JEFATURA (Cobertura y Valorización)
SELECT 
    tipo_producto AS categoria,
    COUNT(codigo_sku_bsale) AS total_skus_migrados,
    SUM(stock_inicial_bodega) AS unidades_totales,
    SUM(stock_inicial_bodega * costo_neto_clp) AS valorizacion_inventario_costo_clp,
    ROUND(AVG((precio_venta_bruto_clp - costo_neto_clp) * 100.0 / NULLIF(precio_venta_bruto_clp, 0)), 1) AS margen_bruto_promedio_pct
FROM vw_matriz_importacion_bsale_xls
GROUP BY tipo_producto
ORDER BY valorizacion_inventario_costo_clp DESC;
