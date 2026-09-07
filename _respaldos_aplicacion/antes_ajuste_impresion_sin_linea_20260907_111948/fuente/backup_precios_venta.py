#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Reemplazo total diario de NPV.dbo.NPVFDPreciosVentaBkp.

La tarea programada conserva, por articulo, el precio con la FECHAINICIO mas
reciente que sea estrictamente anterior al dia de ejecucion.
La aplicacion compara esa instantanea y administra
NPV.dbo.NPVFDPreciosNuevosDiarios al abrir.
"""

import argparse
import datetime as _dt
import json
import os
import re
import sys
import time
import traceback

import pyodbc
from cryptography.fernet import Fernet, InvalidToken


def app_base_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = app_base_path()
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
KEY_PATH = os.path.join(BASE_DIR, "config.key")
LOG_PATH = os.path.join(BASE_DIR, "backup_precios_venta.log")

DATABASE_NAME = "NPV"
BACKUP_TABLE = "NPV.dbo.NPVFDPreciosVentaBkp"
NEW_DAILY_TABLE = "NPV.dbo.NPVFDPreciosNuevosDiarios"
JOB_SQL_SOURCE_TABLE = "NPV.dbo.NPVFDPreciosVenta"


def log(message, exc=None):
    lines = [
        "=" * 80,
        _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        message,
    ]
    if exc is not None:
        lines.append("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)).rstrip())
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    conn_cfg = cfg.setdefault("conexion_odbc", {})
    if conn_cfg.get("password_encrypted") and not conn_cfg.get("password"):
        if not os.path.exists(KEY_PATH):
            raise FileNotFoundError(f"No existe {KEY_PATH}; no se puede descifrar la password.")
        with open(KEY_PATH, "rb") as f:
            key = f.read().strip()
        try:
            conn_cfg["password"] = Fernet(key).decrypt(
                str(conn_cfg["password_encrypted"]).encode("ascii")
            ).decode("utf-8")
        except InvalidToken as exc:
            raise ValueError("config.key no corresponde a config.json.") from exc
    return cfg


def build_connection_string(cfg):
    c = cfg["conexion_odbc"]
    if c.get("dsn"):
        conn_str = f"DSN={c['dsn']};"
    else:
        parts = [f"DRIVER={c.get('driver') or '{ODBC Driver 17 for SQL Server}'}"]
        if c.get("server"):
            parts.append(f"SERVER={c['server']}")
        parts.append(f"DATABASE={c.get('database') or DATABASE_NAME}")
        if c.get("trusted_connection"):
            parts.append("Trusted_Connection=yes")
        else:
            parts.append(f"UID={c.get('username', '')}")
            parts.append(f"PWD={c.get('password', '')}")
        conn_str = ";".join(parts) + ";"

    lower = conn_str.lower()
    if "encrypt=" not in lower:
        conn_str += "Encrypt=no;"
    if "trustservercertificate=" not in lower:
        conn_str += "TrustServerCertificate=yes;"
    return conn_str


def sanitize_connection_string(conn_str):
    return re.sub(r"(PWD=)[^;]*", r"\1***", conn_str or "", flags=re.IGNORECASE)


def connect_with_retry(conn_str, timeout=30, attempts=3, delay_seconds=3):
    for attempt in range(1, attempts + 1):
        try:
            return pyodbc.connect(conn_str, timeout=timeout)
        except pyodbc.Error as exc:
            if attempt >= attempts:
                log("Error conectando a SQL Server.", exc)
                raise
            time.sleep(delay_seconds)


def quote_name(name):
    return "[" + str(name).replace("]", "]]") + "]"


def get_price_column(cursor):
    rows = cursor.execute(
        """
        SELECT COLUMN_NAME
        FROM NPV.INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'dbo'
          AND TABLE_NAME = 'NPVFDPreciosVenta'
          AND COLUMN_NAME IN ('PRECIOVENTA', 'PRECIOSVENTA', 'PRECIO')
        """
    ).fetchall()
    available = {str(row.COLUMN_NAME).upper(): str(row.COLUMN_NAME) for row in rows}
    for candidate in ("PRECIOVENTA", "PRECIOSVENTA", "PRECIO"):
        if candidate in available:
            return available[candidate]
    raise RuntimeError("No se encontro columna de precio en NPV.dbo.NPVFDPreciosVenta.")


def ensure_backup_table(cursor, price_column):
    sql = f"""
    IF OBJECT_ID(N'{BACKUP_TABLE}', N'U') IS NULL
    BEGIN
        SELECT TOP (0)
            pv.ARTICULO,
            pv.{quote_name(price_column)} AS PRECIOVENTA,
            pv.FECHAINICIO,
            pv.FECHAALTA,
            CAST(GETDATE() AS datetime) AS FECHACOPIA
        INTO {BACKUP_TABLE}
        FROM {JOB_SQL_SOURCE_TABLE} AS pv;
    END;

    IF COL_LENGTH(N'{BACKUP_TABLE}', N'FECHAINICIO') IS NULL
    BEGIN
        ALTER TABLE {BACKUP_TABLE}
            ADD FECHAINICIO datetime NULL;
    END;
    """
    cursor.execute(sql)


def ensure_new_daily_table(cursor, price_column):
    """
    Crea la tabla que la aplicacion usa para almacenar los precios detectados.

    La tarea solo asegura su estructura. La aplicacion conserva la
    responsabilidad de comparar, reemplazar y cargar los precios nuevos.
    """
    sql = f"""
    IF OBJECT_ID(N'{NEW_DAILY_TABLE}', N'U') IS NULL
    BEGIN
        SELECT TOP (0)
            CONVERT(date, GETDATE()) AS FECHADETECCION,
            pv.ARTICULO,
            CAST(a.DESCRIPCION AS varchar(255)) AS DESCRIPCION,
            pv.{quote_name(price_column)} AS PRECIOVENTA_ANTERIOR,
            pv.{quote_name(price_column)} AS PRECIOVENTA,
            pv.FECHAINICIO,
            pv.FECHAALTA,
            CAST(GETDATE() AS datetime) AS FECHACOPIA_BKP,
            CAST(GETDATE() AS datetime) AS FECHAREGISTRO
        INTO {NEW_DAILY_TABLE}
        FROM {JOB_SQL_SOURCE_TABLE} AS pv
        LEFT JOIN NPV.dbo.NPVFDArticulos AS a
            ON a.ARTICULO = pv.ARTICULO;

        CREATE INDEX IX_NPVFDPreciosNuevosDiarios_FechaArticulo
            ON {NEW_DAILY_TABLE} (FECHADETECCION, ARTICULO, FECHAINICIO);
    END;
    """
    cursor.execute(sql)


def replace_backup(cursor, price_column):
    sql = f"""
    SET XACT_ABORT ON;

    BEGIN TRANSACTION;

    DELETE FROM {BACKUP_TABLE};

    ;WITH Instantanea AS (
        SELECT
            pv.ARTICULO,
            pv.{quote_name(price_column)} AS PRECIOVENTA,
            pv.FECHAINICIO,
            pv.FECHAALTA,
            ROW_NUMBER() OVER (
                PARTITION BY pv.ARTICULO
                ORDER BY
                    pv.FECHAINICIO DESC,
                    pv.FECHAALTA DESC,
                    pv.{quote_name(price_column)} DESC
            ) AS rn
        FROM {JOB_SQL_SOURCE_TABLE} AS pv
        WHERE pv.FECHAINICIO < CONVERT(date, GETDATE())
    )
    INSERT INTO {BACKUP_TABLE} (
        ARTICULO,
        PRECIOVENTA,
        FECHAINICIO,
        FECHAALTA,
        FECHACOPIA
    )
    SELECT
        pv.ARTICULO,
        pv.PRECIOVENTA,
        pv.FECHAINICIO,
        pv.FECHAALTA,
        GETDATE() AS FECHACOPIA
    FROM Instantanea AS pv
    WHERE pv.rn = 1;

    DECLARE @inserted int = @@ROWCOUNT;

    COMMIT TRANSACTION;

    SELECT @inserted AS inserted_rows;
    """
    cursor.execute(sql)
    while cursor.description is None:
        if not cursor.nextset():
            return None
    row = cursor.fetchone()
    return int(row.inserted_rows) if row else None


def count_source_rows(cursor, price_column):
    return get_validation_counts(cursor, price_column)["copiados"]


def get_validation_counts(cursor, price_column):
    row = cursor.execute(
        f"""
        DECLARE @tomorrow datetime = DATEADD(day, 1, CONVERT(date, GETDATE()));

        ;WITH Instantanea AS (
            SELECT
                pv.ARTICULO,
                ROW_NUMBER() OVER (
                    PARTITION BY pv.ARTICULO
                    ORDER BY
                        pv.FECHAINICIO DESC,
                        pv.FECHAALTA DESC,
                        pv.{quote_name(price_column)} DESC
                ) AS rn
            FROM NPV.dbo.NPVFDPreciosVenta AS pv
            WHERE pv.FECHAINICIO < CONVERT(date, GETDATE())
        ),
        Futuros AS (
            SELECT pv.ARTICULO
            FROM NPV.dbo.NPVFDPreciosVenta AS pv
            WHERE pv.FECHAINICIO >= @tomorrow
        )
        SELECT
            (SELECT COUNT(1) FROM NPV.dbo.NPVFDPreciosVenta) AS precios_total,
            (
                SELECT COUNT(1)
                FROM Instantanea AS i
                WHERE i.rn = 1
            ) AS copiados,
            (
                SELECT COUNT(1)
                FROM Futuros
            ) AS futuros_total,
            (
                SELECT COUNT(1)
                FROM Futuros AS f
                JOIN NPV.dbo.NPVFDArticulos AS a
                    ON a.ARTICULO = f.ARTICULO
            ) AS futuros_con_articulo;
        """
    ).fetchone()

    bkp_exists = cursor.execute(f"SELECT OBJECT_ID(N'{BACKUP_TABLE}', N'U')").fetchval()
    bkp_total = None
    if bkp_exists:
        bkp_total = int(cursor.execute(f"SELECT COUNT(1) FROM {BACKUP_TABLE}").fetchval())

    bkp_has_start_date = False
    if bkp_exists:
        bkp_has_start_date = cursor.execute(
            f"SELECT COL_LENGTH(N'{BACKUP_TABLE}', N'FECHAINICIO')"
        ).fetchval() is not None
    candidates = (
        count_new_daily_candidates(cursor, price_column)
        if bkp_exists and bkp_has_start_date
        else None
    )
    new_daily_exists = bool(
        cursor.execute(f"SELECT OBJECT_ID(N'{NEW_DAILY_TABLE}', N'U')").fetchval()
    )

    return {
        "precios_total": int(row.precios_total or 0),
        "copiados": int(row.copiados or 0),
        "futuros_total": int(row.futuros_total or 0),
        "futuros_con_articulo": int(row.futuros_con_articulo or 0),
        "nuevos_diarios_candidatos": candidates,
        "bkp_total": bkp_total,
        "new_daily_exists": new_daily_exists,
    }


def count_new_daily_candidates(cursor, price_column):
    row = cursor.execute(
        f"""
        ;WITH Cambios AS (
            SELECT
                pv.ARTICULO,
                pv.FECHAINICIO,
                ROW_NUMBER() OVER (
                    PARTITION BY pv.ARTICULO, pv.FECHAINICIO
                    ORDER BY pv.FECHAALTA DESC, pv.{quote_name(price_column)} DESC
                ) AS rn
            FROM {JOB_SQL_SOURCE_TABLE} AS pv
            WHERE pv.FECHAINICIO >= CONVERT(date, GETDATE())
              AND NOT EXISTS (
                SELECT 1
                FROM {BACKUP_TABLE} AS b
                WHERE b.ARTICULO = pv.ARTICULO
                  AND (
                        b.FECHAINICIO = pv.FECHAINICIO
                     OR (b.FECHAINICIO IS NULL AND pv.FECHAINICIO IS NULL)
                  )
            )
        )
        SELECT COUNT(1) AS total
        FROM Cambios
        WHERE rn = 1;
        """
    ).fetchone()
    return int(row.total or 0)


def format_validation_counts(counts):
    parts = [
        f"PreciosVenta total: {counts['precios_total']}",
        f"ultimos precios anteriores a hoy: {counts['copiados']}",
        f"precios futuros: {counts['futuros_total']}",
    ]
    if counts["nuevos_diarios_candidatos"] is None:
        parts.append("comparacion ARTICULO + FECHAINICIO pendiente de migracion")
    else:
        parts.append(f"cambios pendientes contra bkp: {counts['nuevos_diarios_candidatos']}")
    if counts["bkp_total"] is not None:
        parts.append(f"Bkp actual: {counts['bkp_total']}")
    parts.append(
        "NPVFDPreciosNuevosDiarios: "
        + ("existe" if counts["new_daily_exists"] else "no existe")
    )
    return "; ".join(parts) + "."


def main():
    parser = argparse.ArgumentParser(description="Instantanea de precios NPV para detectar cambios al abrir la aplicacion.")
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Guarda el ultimo precio por articulo con FECHAINICIO anterior a hoy.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Valida conexion y cuenta filas sin modificar datos.")
    parser.add_argument("--diagnose", action="store_true", help="Muestra desglose de registros sin modificar datos.")
    args = parser.parse_args()

    cfg = load_config()
    conn_str = build_connection_string(cfg)
    print(sanitize_connection_string(conn_str))

    conn = connect_with_retry(conn_str, timeout=30)
    try:
        cursor = conn.cursor()
        price_column = get_price_column(cursor)

        if args.dry_run or args.diagnose:
            counts = get_validation_counts(cursor, price_column)
            print(f"OK diagnostico. Columna precio fuente: {price_column}.")
            print(format_validation_counts(counts))
            return 0

        ensure_backup_table(cursor, price_column)
        ensure_new_daily_table(cursor, price_column)
        inserted_rows = replace_backup(cursor, price_column)
        conn.commit()
        counts = get_validation_counts(cursor, price_column)
        message = (
            f"Instantanea completada. Columna precio fuente: {price_column}. "
            f"Filas insertadas en bkp: {inserted_rows}. "
            f"{format_validation_counts(counts)}"
        )
        print(message)
        log(message)
        return 0
    except Exception as exc:
        try:
            conn.rollback()
        except Exception as rollback_exc:
            log("Error ejecutando rollback despues de fallo en backup de precios.", rollback_exc)
        log("Error ejecutando backup de precios.", exc)
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
