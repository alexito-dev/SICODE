"""
NotiValores Processor - Motor de simulación para consulta y notificación de Valores SUNAT por RUC.
"""

import asyncio
import random
import re
from datetime import datetime, timedelta
from typing import List, Dict, Any, AsyncGenerator

# Textos repetidos
_BUZ_N_SOL = "Buzón SOL"
_CODIGO = "codigo"
_NOMBRE = "nombre"
_PROGRESS = "progress"
_RAZON_SOCIAL = "razon_social"
_RIESGO = "riesgo"
_SEVERIDAD = "severidad"
_TOTAL = "total"

# Empresa ficticias y reales para simulación realista
RAZONES_SOCIALES_DEMO = {
    "20100047218": "ALICORP S.A.A.",
    "20100130204": "PETROLEOS DEL PERU PETROPERU S.A.",
    "20100070970": "UNION DE CERVECERIAS PERUANAS BACKUS Y JOHNSTON S.A.A.",
    "20100055237": "BANCO DE CREDITO DEL PERU",
    "20100017492": "TELEFONICA DEL PERU S.A.A.",
    "20100070971": "SUPERMERCADOS PERUANOS S.A.",
    "20501234567": "CONSTRUCTORA E INMOBILIARIA DEL PACIFICO S.A.C.",
    "20609876543": "COMERCIALIZADORA LOGISTICA DE LOS ANDES E.I.R.L.",
    "10458796541": "GARCIA LOPEZ JUAN CARLOS",
    "10789654123": "MENDOZA TORRES MARIA ELENA",
}

TIPOS_VALOR = [
    {_CODIGO: "OP", _NOMBRE: "Orden de Pago (OP)", _SEVERIDAD: "alta", _RIESGO: "ALTO"},
    {_CODIGO: "RD", _NOMBRE: "Resolución de Determinación (RD)", _SEVERIDAD: "media", _RIESGO: "MEDIO"},
    {_CODIGO: "RM", _NOMBRE: "Resolución de Multa (RM)", _SEVERIDAD: "media", _RIESGO: "MEDIO"},
    {_CODIGO: "EC", _NOMBRE: "Esquela de Citación / Inconsistencia", _SEVERIDAD: "baja", _RIESGO: "BAJO"},
    {_CODIGO: "CO", _NOMBRE: "Resolución Coactiva - Embargo", _SEVERIDAD: "critica", _RIESGO: "CRÍTICO"},
]

TRIBUTOS = [
    "3081 - IGV CUENTA PROPIA",
    "3031 - RENTA TERCERA CATEGORIA",
    "5210 - ESSALUD REGULAR",
    "5310 - ONP APORTES",
    "3111 - RENTA RETENCIONES 2DA CAT",
    "3051 - RENTA RETENCIONES 4TA CAT",
]

CANALES_NOTIFICACION = {
    "sol": "Buzón Electrónico SOL",
    "email": "Correo Electrónico Notificado",
    "whatsapp": "Mensajería SMS / WhatsApp",
    "cedula": "Cédula Física / Domicilio Fiscal",
}


class NotiValoresProcessor:
    def __init__(self):
        pass

    def validar_ruc(self, ruc: str) -> bool:
        """Valida sintaxis básica de RUC peruano (11 dígitos)."""
        ruc_clean = ruc.strip()
        return bool(re.match(r"^(10|15|17|20)\d{9}$", ruc_clean))

    def obtener_presets(self) -> List[Dict[str, str]]:
        """Retorna lista de RUCs de prueba preconfigurados."""
        return [{"ruc": ruc, _RAZON_SOCIAL: nombre} for ruc, nombre in RAZONES_SOCIALES_DEMO.items()]

    def _generar_valor_aleatorio(self, ruc: str, index: int) -> Dict[str, Any]:
        """Genera un valor SUNAT simulado detallado para un RUC."""
        random.seed(int(ruc) + index)  # NOSONAR (datos demo)
        tipo = random.choice(TIPOS_VALOR)  # NOSONAR (datos demo)
        monto_base = random.randint(850, 125000)  # NOSONAR (datos demo)
        período_mes = random.randint(1, 12)  # NOSONAR (datos demo)
        período_anio = random.randint(2023, 2025)  # NOSONAR (datos demo)
        periodo_str = f"{período_anio}-{período_mes:02d}"

        fecha_emision = datetime.now() - timedelta(days=random.randint(5, 180))  # NOSONAR (datos demo)
        num_valor = f"0{random.randint(10, 99)}-003-{random.randint(1000000, 9999999)}"  # NOSONAR (datos demo)

        razon = RAZONES_SOCIALES_DEMO.get(ruc, f"CONTRIBUYENTE RUC N° {ruc} S.A.C.")

        estados = ["PENDIENTE", "NOTIFICADO", "EN COBRANZA COACTIVA", "IMPUGNADO"]
        estado = random.choice(estados)  # NOSONAR (datos demo)

        return {
            "id": f"VAL-{ruc}-{index + 1}",
            "ruc": ruc,
            _RAZON_SOCIAL: razon,
            "condicion_domicilio": "HABIDO",
            "estado_contribuyente": "ACTIVO",
            "tipo_valor_codigo": tipo[_CODIGO],
            "tipo_valor": tipo[_NOMBRE],
            _RIESGO: tipo[_RIESGO],
            _SEVERIDAD: tipo[_SEVERIDAD],
            "num_documento": num_valor,
            "periodo": periodo_str,
            "tributo": random.choice(TRIBUTOS),  # NOSONAR (datos demo)
            "monto_deuda": float(monto_base),
            "fecha_emision": fecha_emision.strftime("%Y-%m-%d"),
            "fecha_notificacion": (fecha_emision + timedelta(days=2)).strftime("%Y-%m-%d %H:%M"),
            "estado": estado,
            "dependencia": "INTENDENCIA LIMA - MEPECO",
            "canal_notificacion": _BUZ_N_SOL,
            "notificado": True,
        }

    def _separar_rucs(self, rucs_raw: List[str]) -> tuple:
        """Devuelve (RUCs válidos sin repetir, RUCs inválidos)."""
        rucs_validos = []
        rucs_invalidos = []
        for item in rucs_raw:
            cleaned = item.strip()
            if not cleaned:
                continue
            if not self.validar_ruc(cleaned):
                rucs_invalidos.append(cleaned)
            elif cleaned not in rucs_validos:
                rucs_validos.append(cleaned)
        return rucs_validos, rucs_invalidos

    async def _procesar_ruc(self, idx: int, ruc: str, total_rucs: int, canal: str, acumulado: Dict[str, Any]):
        """Simula la consulta de un RUC y emite sus eventos; acumula valores y montos."""
        progress_pct = int((idx / total_rucs) * 100)
        razon = RAZONES_SOCIALES_DEMO.get(ruc, f"EMPRESA CONTRIBUYENTE {ruc}")
        canal_nombre = CANALES_NOTIFICACION.get(canal, _BUZ_N_SOL)
        base = {"step": idx, _TOTAL: total_rucs, _PROGRESS: progress_pct}

        yield {"type": "log", **base, "log": f"🔍 Consultando RUC [{ruc}] - {razon}..."}
        await asyncio.sleep(0.25)

        # Determinar si tiene valores (85% de probabilidad en simulación)
        random.seed(int(ruc))  # NOSONAR (datos demo)
        tiene_valores = random.random() < 0.85  # NOSONAR (datos demo)

        if tiene_valores:
            num_valores = random.randint(1, 3)  # NOSONAR (datos demo)
            acumulado["rucs_con_valores"] += 1
            valores_ruc = []
            for v_idx in range(num_valores):
                val = self._generar_valor_aleatorio(ruc, v_idx)
                val["canal_notificacion"] = canal_nombre
                acumulado["valores"].append(val)
                valores_ruc.append(val)
                acumulado["monto_total"] += val["monto_deuda"]

            deuda = sum(v["monto_deuda"] for v in valores_ruc)
            yield {
                "type": "ruc_found",
                **base,
                "ruc": ruc,
                _RAZON_SOCIAL: razon,
                "valores_count": len(valores_ruc),
                "valores": valores_ruc,
                "log": f"⚠️ RUC [{ruc}] presenta {len(valores_ruc)} Valor(es) SUNAT detectado(s). Deuda acumulada: S/ {deuda:,.2f}",
            }
        else:
            yield {
                "type": "ruc_clean",
                **base,
                "ruc": ruc,
                _RAZON_SOCIAL: razon,
                "log": f"✅ RUC [{ruc}] sin valores ni deudas exigibles registradas a la fecha.",
            }

        await asyncio.sleep(0.2)

        if tiene_valores:
            yield {
                "type": "notification_sent",
                **base,
                "ruc": ruc,
                "canal": canal_nombre,
                "log": f"📩 Notificación enviada exitosamente para RUC [{ruc}] vía {CANALES_NOTIFICACION.get(canal, 'Buzón SOL')}.",
            }
            await asyncio.sleep(0.15)

    async def simular_procesamiento_rucs(
        self, rucs_raw: List[str], canal: str = "sol", incluir_sin_deuda: bool = False
    ) -> AsyncGenerator[Dict[str, Any], None]:
        """
        Simulador asíncrono con Server-Sent Events (SSE).
        Emite progresivamente logs, estados y valores detectados.
        """
        rucs_validos, rucs_invalidos = self._separar_rucs(rucs_raw)
        total_rucs = len(rucs_validos)

        yield {
            "type": "init",
            _TOTAL: total_rucs,
            "invalidos": len(rucs_invalidos),
            "canal_nombre": CANALES_NOTIFICACION.get(canal, _BUZ_N_SOL),
            "timestamp": datetime.now().strftime("%H:%M:%S"),
        }

        if total_rucs == 0:
            yield {
                "type": "complete",
                "summary": {
                    "total_procesados": 0,
                    "rucs_con_valores": 0,
                    "total_valores": 0,
                    "monto_total": 0.0,
                    "notificaciones_enviadas": 0,
                },
                "log": "❌ No se ingresaron RUCs válidos para procesar.",
            }
            return

        acumulado: Dict[str, Any] = {"valores": [], "monto_total": 0.0, "rucs_con_valores": 0}
        for idx, ruc in enumerate(rucs_validos, start=1):
            async for evento in self._procesar_ruc(idx, ruc, total_rucs, canal, acumulado):
                yield evento

        valores = acumulado["valores"]
        monto_total_deuda = acumulado["monto_total"]
        yield {
            "type": "complete",
            _PROGRESS: 100,
            "summary": {
                "total_procesados": total_rucs,
                "rucs_con_valores": acumulado["rucs_con_valores"],
                "total_valores": len(valores),
                "monto_total": round(monto_total_deuda, 2),
                "notificaciones_enviadas": len(valores),
            },
            "valores_todos": valores,
            "log": f"🎉 Simulación finalizada. {len(valores)} valores notificados en total por S/ {monto_total_deuda:,.2f}.",
        }

    def generar_html_cedula(self, valor_data: Dict[str, Any]) -> str:
        """Genera plantilla HTML formal de la Cédula de Notificación de Valor SUNAT."""
        fecha_hoy = datetime.now().strftime("%d/%m/%Y")
        hora_hoy = datetime.now().strftime("%H:%M:%S")

        return f"""
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <title>Cédula de Notificación N° {valor_data.get("num_documento")}</title>
            <style>
                body {{ font-family: Arial, sans-serif; font-size: 13px; color: #1a1a1a; padding: 25px; background: #fff; }}
                .header {{ border-bottom: 2px solid #005691; padding-bottom: 15px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center; }}
                .logo-title {{ font-size: 18px; font-weight: bold; color: #005691; }}
                .sub-header {{ font-size: 11px; color: #666; }}
                .box {{ border: 1px solid #cbd5e1; border-radius: 6px; padding: 15px; margin-bottom: 15px; background: #f8fafc; }}
                .title-doc {{ text-align: center; font-size: 15px; font-weight: bold; text-transform: uppercase; margin: 15px 0; color: #0f172a; text-decoration: underline; }}
                table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
                th, td {{ border: 1px solid #94a3b8; padding: 8px 12px; text-align: left; font-size: 12px; }}
                th {{ background-color: #e2e8f0; color: #1e293b; font-weight: bold; }}
                .badge {{ display: inline-block; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 11px; color: white; }}
                .badge-danger {{ background-color: #ef4444; }}
                .badge-warning {{ background-color: #f59e0b; }}
                .badge-info {{ background-color: #3b82f6; }}
                .footer {{ margin-top: 30px; font-size: 10px; color: #64748b; text-align: center; border-top: 1px dashed #cbd5e1; padding-top: 10px; }}
                .barcode {{ text-align: center; font-family: 'Courier New', Courier, monospace; letter-spacing: 4px; margin-top: 15px; font-weight: bold; }}
            </style>
        </head>
        <body>
            <div class="header">
                <div>
                    <div class="logo-title">SUNAT - SUPERINTENDENCIA NACIONAL DE ADUANAS Y DE ADMINISTRACIÓN TRIBUTARIA</div>
                    <div class="sub-header">{valor_data.get("dependencia", "INTENDENCIA LIMA MEPECO")}</div>
                </div>
                <div style="text-align: right;">
                    <strong>CONSTANCIA DE NOTIFICACIÓN ELECTRONICA</strong><br>
                    <span style="font-size: 11px; color: #475569;">Fecha/Hora: {fecha_hoy} {hora_hoy}</span>
                </div>
            </div>

            <div class="title-doc">CÉDULA DE NOTIFICACIÓN DE VALOR TRIBUTARIO</div>

            <div class="box">
                <strong>DATOS DEL CONTRIBUYENTE NOTIFICADO:</strong><br>
                <table>
                    <tr>
                        <td style="width: 25%;"><strong>RUC:</strong></td>
                        <td>{valor_data.get("ruc")}</td>
                        <td style="width: 25%;"><strong>Condición Domicilio:</strong></td>
                        <td><span style="color: green; font-weight: bold;">{valor_data.get("condicion_domicilio")}</span></td>
                    </tr>
                    <tr>
                        <td><strong>Razón Social / Nombre:</strong></td>
                        <td colspan="3"><strong>{valor_data.get("razon_social")}</strong></td>
                    </tr>
                    <tr>
                        <td><strong>Estado Contribuyente:</strong></td>
                        <td>{valor_data.get("estado_contribuyente")}</td>
                        <td><strong>Canal Notificación:</strong></td>
                        <td>{valor_data.get("canal_notificacion")}</td>
                    </tr>
                </table>
            </div>

            <div class="box" style="background: #fff;">
                <strong>DETALLE DEL VALOR DEUDA TRIBUTARIA:</strong>
                <table>
                    <thead>
                        <tr>
                            <th>N° Documento Valor</th>
                            <th>Tipo de Valor</th>
                            <th>Período</th>
                            <th>Tributo / Concepto</th>
                            <th>Monto Deuda S/</th>
                            <th>Estado Actual</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td><strong>{valor_data.get("num_documento")}</strong></td>
                            <td>{valor_data.get("tipo_valor")}</td>
                            <td>{valor_data.get("periodo")}</td>
                            <td>{valor_data.get("tributo")}</td>
                            <td style="color: #b91c1c; font-weight: bold; font-size: 13px;">S/ {valor_data.get("monto_deuda", 0):,.2f}</td>
                            <td><span class="badge badge-danger">{valor_data.get("estado")}</span></td>
                        </tr>
                    </tbody>
                </table>
            </div>

            <div class="box">
                <strong>FUNDAMENTO Y BASE LEGAL DE LA NOTIFICACIÓN:</strong><br>
                <p style="margin-top: 6px; line-height: 1.4; color: #334155;">
                    Se notifica al contribuyente la existencia de la acreencia tributaria detallada en el presente documento,
                    de conformidad con los Artículos 104 y 115 del Código Tributario (T.U.O. del D.S. N° 133-2013-EF).
                    El plazo para la subsanación, pago o cancelación amigable de la obligación tributaria es de cinco (05) días hábiles
                    contados a partir del día siguiente de la presente notificación electrónica.
                </p>
            </div>

            <div class="barcode">
                ||||| |||||||||| |||||||| |||||||||||||| ||||||<br>
                REF-{valor_data.get("id")}-SUNAT-NOTI
            </div>

            <div class="footer">
                Documento generado por el Sistema SICODE - Módulo NotiValores SUNAT (Simulación Operativa).<br>
                Código de Autenticidad Digital: {hash(valor_data.get("id", "123"))} - Firma Digital Emitida.
            </div>
        </body>
        </html>
        """
