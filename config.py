"""
Archivo de Configuración Central para el proyecto.

Aquí se definen las reglas de negocio, como el orden de los slides
y el mapeo de columnas para los criterios de evaluación.
"""

# --- ORDEN DE SLIDES (SUBDOMINIOS) ---
# Define el orden de los slides .txt dentro de cada PPTX de dominio.
CUSTOM_SORT_ORDER = {
    "canales": [
        "sitio_publico",
        "web_retail",
        "app_retail",
        "web_empresas",
        "app_empresas",
        "call_center",
        "sucursales",
        "wallet",
        "plataformas_de_terceros",
        "atm",
        "pos",
        "gestion_de_contenidos",
        "crm",
        "martech",
        "otros",
    ],
    "IntegracionProcesos": [
        "api_management_gateway",
        "microservicios",
        "event_broker",
        "bpm",
        "brms",
        "message_broker",
        "esb_eai",
        "transferencia_de_archivos",
        "rpa",
    ],
    "soporteempresarial": [
        "erp",
        "hr",
        "auditoria",
        "administracion_de_contratos",
        "riesgo",
        "cumplimiento",
        "mercados_tesoreria_comex",
    ],
    "corebanking": [
        "depositos_cuentas_inversiones",
        "creditos_e_inversiones",
        "cuentas_bancarias_comercio_minorista_y_empresas",
        "pagos_y_tarjetas",
        "cobranzas",
    ],
    "datos": [
        "almacenamiento",
        "consumo",
        "gobierno",
        "integracion_ingestion_y_procesamiento",
        "servicios_datos",
    ],
    "operacionti": [
        "planeacion_documentacion_y_diseno",
        "desarrollo",
        "pruebas",
        "despliegue",
        "monitoreo_y_operaciones",
    ],
}


# --- MAPEADO DE CRITERIOS Y COLUMNAS DE LA BD ---
# Define qué columna de la BD corresponde a cada criterio de evaluación.
# El generador de slides usará este mapa para leer los datos.
CRITERIA_DB_MAP = {
    # Criterios de Evaluación
    "obsolescencia": "nivel_de_obsolescencia_1",
    "escalabilidad": "tiene_alta_disponibilidad_1",
    "estabilidad": "ha_presentado_caidas_o_degradacion_del_servicio_en_los_ultimo_1",
    "extensibilidad": "extensibilidad",
    "seguridad": "seguridad",
    "ux": "ux",
    # Criterio 'agilidad' depende de dos columnas
    "agilidad": ("devops_1", "despliegue_a_pdn_automatizado_1"),
    # Criterios Fijos (no dependen de la BD)
    "acople": None,  # Siempre es "Parcialmente"
    "cobertura": None,  # Siempre está vacío
    # Iconos de la fila
    "icon_sas": "sas",
    "icon_cots": "nivel_de_customizacion",
    "icon_cloud": "nube_vs_onpremise",
    "icon_regional": "bns",  # Usa la misma columna que extensibilidad
    # Tecnología
    "tecnologia": "tecnologia_subyacente",
}
