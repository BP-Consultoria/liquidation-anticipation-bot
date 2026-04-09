import pyodbc
import pandas as pd
from config import settings


class DatabaseService:
    def __init__(self):
        self._conn = None

    def conectar(self):
        conn_str = (
            f"DRIVER={{{settings.db_odbc_driver}}};"
            f"SERVER={settings.db_host},{settings.db_port};"
            f"DATABASE={settings.db_name};"
            f"UID={settings.db_user};"
            f"PWD={settings.db_password}"
        )
        self._conn = pyodbc.connect(conn_str)
        print(f"[DB] Conectado ao banco {settings.db_name}")
        return self._conn

    def desconectar(self):
        if self._conn:
            self._conn.close()
            self._conn = None
            print("[DB] Conexão encerrada")

    def buscar_antecipacoes_hoje(self):
        if not self._conn:
            raise ConnectionError("Conexão não estabelecida. Chame conectar() primeiro.")

        cursor = self._conn.cursor()
        cursor.execute(
            "SELECT "
            "    a.Bordero, a.Cedente, a.Titulo, a.Valor, "
            "    a.Is_inserted, a.Control_id, a.Is_send, "
            "    a.Titulo_Completo, a.created_at, a.Vencimento, "
            "    a.Valor_Desagio, b.Valor_Liquido_Final "
            "FROM anticipation_db.dbo.anticipation_db AS a "
            "INNER JOIN anticipation_db.dbo.borderos AS b "
            "    ON a.Bordero = b.Bordero "
            "WHERE "
            "    a.Is_inserted = 1 "
            "    AND a.Control_id = 4 "
            "    AND a.created_at >= CAST(GETDATE() AS DATE) "
            "    AND a.created_at < DATEADD(DAY, 1, CAST(GETDATE() AS DATE))"
        )
        colunas = [col[0] for col in cursor.description]
        linhas = cursor.fetchall()
        return [dict(zip(colunas, row)) for row in linhas]

    def buscar_dados_para_rpa(self) -> pd.DataFrame:
        if not self._conn:
            raise ConnectionError("Conexão não estabelecida. Chame conectar() primeiro.")

        query = (
            "SELECT "
            "    a.Bordero, a.Cedente, a.Titulo, a.Valor, "
            "    a.created_at, a.Vencimento, a.Nome_Cedente, "
            "    a.Valor_Total_Desagio, b.Valor_Liquido_Final "
            "FROM anticipation_db.dbo.anticipation_db AS a "
            "INNER JOIN anticipation_db.dbo.borderos AS b "
            "    ON a.Bordero = b.Bordero "
            "WHERE "
            "    a.Is_inserted = 1 "
            "    AND a.Control_id = 4 "
            "    AND a.created_at >= CAST(GETDATE() AS DATE) "
            "    AND a.created_at < DATEADD(DAY, 1, CAST(GETDATE() AS DATE))"
        )
        return pd.read_sql(query, self._conn)

    def atualizar_valor_liquido(self, numero_bordero: int, valor: float):
        if not self._conn:
            raise ConnectionError("Conexão não estabelecida. Chame conectar() primeiro.")

        cursor = self._conn.cursor()
        cursor.execute(
            "UPDATE [anticipation_db].[dbo].[borderos] "
            "SET Valor_Liquido_Final = ? "
            "WHERE Bordero = ?",
            valor, numero_bordero,
        )
        self._conn.commit()
        print(f"[DB] Bordero {numero_bordero} atualizado com valor R$ {valor:.2f}")


db_service = DatabaseService()
