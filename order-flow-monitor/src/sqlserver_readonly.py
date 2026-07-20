#!/usr/bin/env python3
"""
SQL Server 只读查询模块（强制安全版）
所有 SQL Server 查询必须经过此模块，六层安全过滤不可绕过。

用法:
    from sqlserver_readonly import query_to_excel
    result = query_to_excel(sql="SELECT * FROM Orders")
"""

import pyodbc
import pandas as pd
import re
import os
import sys
import json
from datetime import datetime
from pathlib import Path

# ============ 配置 ============
CONFIG_PATH = os.path.expanduser("~/.config/sqlserver_readonly/config.json")
ENCRYPTED_CONFIG_PATH = os.path.expanduser("~/.config/sqlserver_readonly/config.enc")
KEY_PATH = os.path.expanduser("~/.config/sqlserver_readonly/.key")
OUTPUT_DIR = "/mnt/d/SQL导出"

# ============ 危险操作黑名单 ============
FORBIDDEN_KEYWORDS = [
    'CREATE', 'ALTER', 'DROP', 'TRUNCATE', 'RENAME',
    'INSERT', 'UPDATE', 'DELETE', 'MERGE', 'UPSERT',
    'GRANT', 'REVOKE', 'DENY',
    'COMMIT', 'ROLLBACK', 'SAVEPOINT', 'BEGIN TRANSACTION',
    'EXEC', 'EXECUTE', 'SP_', 'XP_', 'SYS.', 'INFORMATION_SCHEMA',
    'DBCC', 'BACKUP', 'RESTORE', 'SHUTDOWN', 'KILL',
    'BULK INSERT', 'OPENROWSET', 'OPENDATASOURCE',
    'INTO OUTFILE', 'INTO DUMPFILE', 'LOAD_FILE',
]

FORBIDDEN_CLAUSES = ['INTO', 'UNION', 'UNION ALL']

# ============ 安全过滤器 ============
class SecurityError(Exception):
    pass

def _validate_sql(sql: str) -> None:
    """六层安全过滤，失败直接抛 SecurityError"""
    if not sql or not isinstance(sql, str):
        raise SecurityError("SQL 不能为空")

    upper = sql.upper().strip()

    # 第1层：前缀白名单
    if not upper.startswith('SELECT'):
        raise SecurityError("只允许 SELECT 查询")

    # 第2层：关键字黑名单（整词匹配，避免子串误杀）
    import re
    forbidden_pattern = re.compile(
        r'\b(' + '|'.join(re.escape(kw) for kw in FORBIDDEN_KEYWORDS) + r')\b'
    )
    if forbidden_pattern.search(upper):
        matched = forbidden_pattern.search(upper).group(1)
        raise SecurityError(f"检测到危险关键字: {matched}")

    # 第3层：子句黑名单
    for clause in FORBIDDEN_CLAUSES:
        if clause in upper:
            raise SecurityError(f"检测到禁止子句: {clause}")

    # 第4层：符号过滤（防多语句）
    if ';' in sql:
        raise SecurityError("禁止分号，不允许执行多条语句")

    # 第5层：UNION 显式过滤
    if 'UNION' in upper:
        raise SecurityError("禁止 UNION 操作")

    # 第6层：长度限制
    if len(sql) > 10000:
        raise SecurityError("SQL 长度超过限制（10000字符）")

# ============ 配置管理 ============
def _load_config() -> dict:
    """从配置文件读取连接信息（优先读取加密配置）"""
    # 优先读取加密配置
    if os.path.exists(ENCRYPTED_CONFIG_PATH) and os.path.exists(KEY_PATH):
        try:
            from cryptography.fernet import Fernet
            import hashlib
            import base64

            with open(KEY_PATH, 'r') as f:
                key_hex = f.read().strip()
            key_bytes = hashlib.sha256(key_hex.encode()).digest()
            key_b64 = base64.urlsafe_b64encode(key_bytes)
            fernet = Fernet(key_b64)

            with open(ENCRYPTED_CONFIG_PATH, 'r') as f:
                cfg = json.load(f)

            # 解密密码
            encrypted_pw = cfg.get('encrypted_password')
            if encrypted_pw:
                cfg['password'] = fernet.decrypt(encrypted_pw.encode()).decode()

            return cfg
        except Exception as e:
            raise RuntimeError(f"读取加密配置失败: {e}")

    # 回退到明文配置（兼容旧方式）
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)

    raise FileNotFoundError(
        f"配置文件不存在: {ENCRYPTED_CONFIG_PATH} 或 {CONFIG_PATH}\n"
        "请先创建配置文件"
    )

def _build_connection_string(cfg: dict) -> str:
    """构建 ODBC 连接字符串"""
    return (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['username']};"
        f"PWD={cfg['password']};"
        f"TrustServerCertificate=yes;"
        f"Encrypt=yes"
    )

# ============ 核心查询接口 ============
def query_to_excel(
    sql: str,
    output_name: str = None,
    output_dir: str = OUTPUT_DIR,
    max_rows: int = 100000
) -> dict:
    """
    执行只读查询并导出 Excel。

    Args:
        sql: SELECT 查询语句（必须经过安全过滤）
        output_name: 输出文件名（默认自动生成）
        output_dir: 输出目录（默认 /mnt/d/SQL导出）
        max_rows: 最大返回行数

    Returns:
        {"success": True, "file": "绝对路径", "rows": 行数, "columns": 列数, "time_ms": 耗时}

    Raises:
        SecurityError: SQL 未通过安全过滤
        pyodbc.Error: 数据库连接或查询错误
    """
    # 强制安全过滤（不可绕过）
    _validate_sql(sql)

    # 加载配置
    cfg = _load_config()
    conn_str = _build_connection_string(cfg)

    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)

    # 自动生成文件名
    if not output_name:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = f"query_{timestamp}.xlsx"
    if not output_name.endswith('.xlsx'):
        output_name += '.xlsx'
    output_path = os.path.join(output_dir, output_name)

    # 执行查询
    start = datetime.now()
    conn = pyodbc.connect(conn_str, timeout=30)
    try:
        df = pd.read_sql(sql, conn)
        if len(df) > max_rows:
            raise SecurityError(f"返回行数 {len(df)} 超过限制 {max_rows}")
    finally:
        conn.close()

    elapsed = int((datetime.now() - start).total_seconds() * 1000)

    # 写入 Excel
    df.to_excel(output_path, index=False, engine='openpyxl')

    return {
        "success": True,
        "file": output_path,
        "rows": len(df),
        "columns": len(df.columns),
        "time_ms": elapsed
    }

# ============ 表结构读取接口 ============
def get_tables() -> list:
    """
    获取数据库中所有用户表名。
    
    Returns:
        ["TableName1", "TableName2", ...]
    """
    cfg = _load_config()
    conn_str = _build_connection_string(cfg)
    conn = pyodbc.connect(conn_str, timeout=30)
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_SCHEMA = 'dbo'
            ORDER BY TABLE_NAME
        """)
        return [row[0] for row in cursor.fetchall()]
    finally:
        conn.close()

def get_columns(table_name: str) -> list:
    """
    获取指定表的所有列信息。
    
    Args:
        table_name: 表名（如 "T_BOM_Item"）
        
    Returns:
        [{"name": "Id", "type": "int", "max_length": 4, "nullable": False}, ...]
    """
    # 防止注入：只允许字母数字下划线
    if not re.match(r'^[A-Za-z0-9_]+$', table_name):
        raise SecurityError(f"非法表名: {table_name}")
    
    cfg = _load_config()
    conn_str = _build_connection_string(cfg)
    conn = pyodbc.connect(conn_str, timeout=30)
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT 
                COLUMN_NAME,
                DATA_TYPE,
                CHARACTER_MAXIMUM_LENGTH,
                IS_NULLABLE
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = ? AND TABLE_SCHEMA = 'dbo'
            ORDER BY ORDINAL_POSITION
        """, (table_name,))
        
        columns = []
        for row in cursor.fetchall():
            columns.append({
                "name": row[0],
                "type": row[1],
                "max_length": row[2],
                "nullable": row[3] == "YES"
            })
        return columns
    finally:
        conn.close()

def get_table_preview(table_name: str, limit: int = 5) -> dict:
    """
    获取表的前 N 行数据（用于快速查看）。
    
    Args:
        table_name: 表名
        limit: 预览行数（默认 5）
        
    Returns:
        {"columns": [...], "rows": [[...], ...]}
    """
    if not re.match(r'^[A-Za-z0-9_]+$', table_name):
        raise SecurityError(f"非法表名: {table_name}")
    if limit > 100:
        raise SecurityError("预览行数不能超过 100")
    
    cfg = _load_config()
    conn_str = _build_connection_string(cfg)
    conn = pyodbc.connect(conn_str, timeout=30)
    try:
        cursor = conn.cursor()
        cursor.execute(f"SELECT TOP {limit} * FROM [{table_name}]")
        
        columns = [desc[0] for desc in cursor.description]
        rows = [list(row) for row in cursor.fetchall()]
        
        return {"columns": columns, "rows": rows}
    finally:
        conn.close()

# ============ CLI 入口 ============
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="SQL Server 只读查询")
    parser.add_argument("--sql", required=True, help="SELECT 查询语句")
    parser.add_argument("--output", default=None, help="输出文件名")
    parser.add_argument("--dir", default=OUTPUT_DIR, help="输出目录")
    args = parser.parse_args()

    try:
        result = query_to_excel(sql=args.sql, output_name=args.output, output_dir=args.dir)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except SecurityError as e:
        print(json.dumps({"success": False, "error": f"安全拦截: {e}"}, ensure_ascii=False))
        sys.exit(1)
    except Exception as e:
        print(json.dumps({"success": False, "error": str(e)}, ensure_ascii=False))
        sys.exit(1)
