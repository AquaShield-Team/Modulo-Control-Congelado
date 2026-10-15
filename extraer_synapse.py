import os
import sys
import pyodbc
import struct
import json
import pandas as pd
from datetime import datetime

# ==============================================================================
# AQUASHIELD - CONECTOR AUTOMATICO AZURE SYNAPSE ANALYTICS (DATA LAKE)
# ==============================================================================

SERVER = 'synw-prd-datalake-eus2-ondemand.sql.azuresynapse.net'
DATABASE = 'ondemand'
DRIVER = '{ODBC Driver 18 for SQL Server}'

SQL_QUERY = """
WITH PedidosCongelados AS (
    -- Consulta 1: Datos de Venta (COM.ZPR_SD_REPVTA01)
    SELECT [Vbeln]
          ,[Kunnr]
          ,[Name1]
          ,[NomDmercaderia]
          ,[Bstkd]
          ,[Bstdk]
          ,[Estaprod]
          ,[Vkorg]
          ,[Zlschtxt]
          ,[Inco1]
          ,[Bl]
          ,[NumCont]
          ,[TxtTransportista]
          ,[Nave]
          ,[PuertoDes]
          ,[Fkfolio]
          ,[Fkvbeln]
          ,ROW_NUMBER() OVER (
              PARTITION BY [Vbeln] 
              ORDER BY 
                  -- Prioridad 1: Si ya está facturado en alguna línea, conservar esa fila
                  CASE WHEN [Fkvbeln] IS NOT NULL AND [Fkvbeln] <> '' AND [Fkvbeln] <> '0' THEN 0 ELSE 1 END ASC,
                  -- Prioridad 2: Conservar la fila que contenga el contenedor
                  CASE WHEN [NumCont] IS NOT NULL AND [NumCont] <> '' THEN 0 ELSE 1 END ASC,
                  -- Prioridad 3: Conservar la fila que contenga BL
                  CASE WHEN [Bl] IS NOT NULL AND [Bl] <> '' THEN 0 ELSE 1 END ASC,
                  -- Prioridad 4: Conservar la fila que contenga Contrato / OC
                  CASE WHEN [Bstkd] IS NOT NULL AND [Bstkd] <> '' THEN 0 ELSE 1 END ASC,
                  [Id] DESC
          ) AS Fila
      FROM [COM].[ZPR_SD_REPVTA01]
      WHERE [Estaprod] = 'CONGELADO'
        AND ISNULL([TxtTransportista], '') NOT LIKE '%TERRESTRE%'
        -- Regla: Solo pedidos marítimos que comiencen con 40
        AND [Vbeln] LIKE '40%'
        -- Regla de Temporada Oficial: A partir de Junio 2026 (pedidos >= 40566606)
        AND [Vbeln] >= '40566606'
),
ContenedoresAgrupados AS (
    -- Agrupación de todos los contenedores distintos por pedido
    SELECT [Vbeln],
           STRING_AGG([NumCont], ' / ') AS [TodosContenedores]
    FROM (
        SELECT DISTINCT [Vbeln], LTRIM(RTRIM([NumCont])) AS [NumCont]
        FROM [COM].[ZPR_SD_REPVTA01]
        WHERE [NumCont] IS NOT NULL AND LTRIM(RTRIM([NumCont])) <> ''
          AND [Estaprod] = 'CONGELADO'
          AND [Vbeln] >= '40566606'
    ) C
    GROUP BY [Vbeln]
),
PedidosUnicos AS (
    -- Consulta 2: Datos de Transporte en Vivo (Project44)
    SELECT [pedido]
          ,[containerId]
          ,[etd]
          ,[etaPredictionP44]
          ,ROW_NUMBER() OVER (
              PARTITION BY [pedido] 
              ORDER BY 
                  CASE WHEN [etaPredictionP44] IS NOT NULL THEN 0 ELSE 1 END ASC,
                  [etd] ASC
          ) AS Fila
      FROM [COM].[project44]
),
FacturasFlujo AS (
    -- Consulta 3: Facturación Real de Exportaciones (AQ y BK) para recuperar facturas ausentes en REPVTA01
    SELECT [Vbelv], [Factura], [Xblnr], [Fkdat],
           ROW_NUMBER() OVER (PARTITION BY [Vbelv] ORDER BY [Fkdat] DESC, [Factura] DESC) as Fila
    FROM (
        SELECT [Vbelv], [Factura], [Xblnr], [Fkdat] FROM [COM].[ZPR_SD_RP_FLUJOEX_AQ] WHERE [Factura] IS NOT NULL AND [Factura] <> ''
        UNION ALL
        SELECT [Vbelv], [Factura], [Xblnr], [Fkdat] FROM [COM].[ZPR_SD_RP_FLUJOEX_BK] WHERE [Factura] IS NOT NULL AND [Factura] <> ''
    ) F
)
-- UNIÓN Y RESULTADO FINAL EN VIVO
SELECT TOP (100000) 
       PC.[Vbeln]                                   AS [Vbeln]
      ,PC.[Kunnr]                                   AS [Codigo Cliente]
      ,PC.[Name1]                                   AS [Nombre Cliente]
      ,PC.[NomDmercaderia]                          AS [Destinatario]
      ,PC.[Bstkd]                                   AS [Nro Contrato]
      ,FORMAT(TRY_CAST(PC.[Bstdk] AS DATE), 'dd-MM-yyyy') AS [Fecha Contrato]
      ,CASE 
          WHEN PC.[Vkorg] LIKE '%S170%' THEN 'S170'
          WHEN PC.[Vkorg] LIKE '%S070%' THEN 'S070'
          ELSE PC.[Vkorg]
       END                                          AS [Sociedad]
      ,CASE 
          WHEN UPPER(PC.[Zlschtxt]) LIKE '%CRED%' OR UPPER(PC.[Zlschtxt]) LIKE '%CTA CTE%' THEN 'Crédito'
          WHEN UPPER(PC.[Zlschtxt]) LIKE '%ANTIC%' THEN 'Anticipo'
          WHEN UPPER(PC.[Zlschtxt]) LIKE '%TRANSF%' THEN 'Transferencia'
          WHEN UPPER(PC.[Zlschtxt]) LIKE '%CARTA%' THEN 'Carta de Crédito'
          WHEN UPPER(PC.[Zlschtxt]) LIKE '%COPIA%' THEN 'Cobranza'
          ELSE ISNULL(PC.[Zlschtxt], '')
       END                                          AS [Condicion Venta]
      ,PC.[Inco1]                                   AS [Incoterm]
      ,LTRIM(RTRIM(PC.[Bl]))                        AS [Nro BL]
      
      -- Contenedor: Prioriza NumCont agrupado de REPVTA01, fallback a Project44
      ,CASE 
          WHEN CA.[TodosContenedores] IS NOT NULL AND CA.[TodosContenedores] <> '' THEN CA.[TodosContenedores]
          WHEN PC.[NumCont] IS NOT NULL AND PC.[NumCont] <> '' THEN LTRIM(RTRIM(PC.[NumCont]))
          ELSE PU.[containerId]
       END                                          AS [Nro Contenedor]
       
      ,PC.[TxtTransportista]                        AS [Naviera]
      ,PC.[Nave]                                    AS [Nave]
      
      -- Puerto Destino: limpia código numérico inicial (ej: 000511) y sufijo ', PUERTO'
      ,CASE 
          WHEN PC.[PuertoDes] IS NULL OR PC.[PuertoDes] = '' THEN ''
          WHEN PATINDEX('%[^0-9 ]%', PC.[PuertoDes]) = 0 THEN ''
          ELSE LTRIM(RTRIM(
              REPLACE(
              REPLACE(
              REPLACE(
              REPLACE(
              REPLACE(
                  SUBSTRING(PC.[PuertoDes], PATINDEX('%[^0-9 ]%', PC.[PuertoDes]), LEN(PC.[PuertoDes])),
              ', PUERTO.', ''),
              ', PUERTO', ''),
              ',PUERTO', ''),
              ', PUERT', ''),
              '(PUERTO)', '')
          ))
       END                                          AS [Puerto Destino]
      
      -- Factura SAP (match exacto con VF05N)
      ,CASE 
          WHEN TRY_CAST(PC.[Fkvbeln] AS BIGINT) IS NOT NULL AND TRY_CAST(PC.[Fkvbeln] AS BIGINT) <> 0 
               THEN CAST(TRY_CAST(PC.[Fkvbeln] AS BIGINT) AS VARCHAR(50))
          WHEN TRY_CAST(FF.[Factura] AS BIGINT) IS NOT NULL AND TRY_CAST(FF.[Factura] AS BIGINT) <> 0 
               THEN CAST(TRY_CAST(FF.[Factura] AS BIGINT) AS VARCHAR(50))
          ELSE 'No facturado'
       END                                          AS [Folio SAP]

      -- Folio DTE / Factura Legal SII
      ,CASE 
          WHEN TRY_CAST(PC.[Fkfolio] AS BIGINT) IS NOT NULL AND TRY_CAST(PC.[Fkfolio] AS BIGINT) <> 0 
               THEN CAST(TRY_CAST(PC.[Fkfolio] AS BIGINT) AS VARCHAR(50))
          WHEN TRY_CAST(FF.[Xblnr] AS BIGINT) IS NOT NULL AND TRY_CAST(FF.[Xblnr] AS BIGINT) <> 0 
               THEN CAST(TRY_CAST(FF.[Xblnr] AS BIGINT) AS VARCHAR(50))
          ELSE 'No facturado'
       END                                          AS [Nro Factura]
       
      -- Fechas de Project44 en Vivo
      ,FORMAT(TRY_CAST(PU.[etd] AS DATETIME), 'dd-MM-yyyy') AS [ETD]
      ,FORMAT(TRY_CAST(PU.[etaPredictionP44] AS DATETIME), 'dd-MM-yyyy') AS [ETA]
      
FROM PedidosCongelados PC
LEFT JOIN ContenedoresAgrupados CA
    ON PC.[Vbeln] = CA.[Vbeln]
LEFT JOIN PedidosUnicos PU 
    ON PC.[Vbeln] = PU.[pedido] AND PU.Fila = 1
LEFT JOIN FacturasFlujo FF
    ON PC.[Vbeln] = FF.[Vbelv] AND FF.Fila = 1
WHERE PC.Fila = 1
ORDER BY PC.[Vbeln] DESC;
"""

def get_azure_token():
    """Obtiene token de Azure Entra ID para la base de datos"""
    try:
        from azure.identity import InteractiveBrowserCredential
        cred = InteractiveBrowserCredential()
        token = cred.get_token("https://database.windows.net//.default")
        return token.token
    except Exception as e:
        print(f"[AVISO] No se pudo obtener token con InteractiveBrowserCredential: {e}")
        return None

USER_EMAIL = 'marcelo.ramirez@aquachile.com'

def connect_synapse():
    print(f"-> Conectando a Azure Synapse: {SERVER}...")
    
    # Intento 1: Conexión Integrada con cuenta Windows / AquaChile SSO (100% silenciosa y automática)
    try:
        conn_str = (
            f"DRIVER={DRIVER};"
            f"SERVER={SERVER};"
            f"DATABASE={DATABASE};"
            "Authentication=ActiveDirectoryIntegrated;"
            "Encrypt=yes;"
            "TrustServerCertificate=yes;"
            "LoginTimeout=15;"
        )
        cnxn = pyodbc.connect(conn_str)
        print(f"   [OK] Conexión establecida automáticamente vía ActiveDirectoryIntegrated (SSO {USER_EMAIL}).")
        return cnxn
    except Exception as e1:
        print(f"   [INFO] Intento 1 (ActiveDirectoryIntegrated): {str(e1)[:120]}")

    # Intento 2: Conexión Interactiva pasando UID directamente para no pedir escribir/seleccionar correo
    try:
        conn_str = (
            f"DRIVER={DRIVER};"
            f"SERVER={SERVER};"
            f"DATABASE={DATABASE};"
            "Authentication=ActiveDirectoryInteractive;"
            f"UID={USER_EMAIL};"
            "Encrypt=yes;"
            "TrustServerCertificate=yes;"
            "LoginTimeout=30;"
        )
        cnxn = pyodbc.connect(conn_str)
        print("   [OK] Conexión establecida vía ActiveDirectoryInteractive.")
        return cnxn
    except Exception as e2:
        print(f"   [INFO] Intento 2 (ActiveDirectoryInteractive): {str(e2)[:120]}")

    return None

def main():
    print("==================================================================")
    print("   AQUASHIELD - EXTRACCION AUTOMATICA DESDE AZURE SYNAPSE")
    print("==================================================================")
    
    cnxn = connect_synapse()
    if not cnxn:
        print("\n[ERROR] No fue posible conectar directamente a Azure Synapse.")
        print("-> Verificando si existe archivo exportado pedidos_sql.xlsx...")
        if os.path.exists("pedidos_sql.xlsx") or os.path.exists("pedidos_sql.csv"):
            print("   [OK] Archivo exportado encontrado. Se procederá con la triangulación local.")
            return True
        else:
            print("   [INFO] Por favor ejecuta la consulta en SSMS y guárdala como 'pedidos_sql.xlsx' en esta carpeta.")
            return False

    print("-> Ejecutando consulta SQL en el Data Lake...")
    df = pd.read_sql(SQL_QUERY, cnxn)
    cnxn.close()
    print(f"   [OK] Se descargaron exitosamente {len(df)} pedidos desde Synapse!")

    output_file = "pedidos_sql.xlsx"
    df.to_excel(output_file, index=False)
    print(f"-> Guardado en {output_file} listo para triangulación.")
    return True

if __name__ == '__main__':
    ok = main()
    sys.exit(0 if ok else 1)
