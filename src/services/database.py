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
            "    b.Valor_Liquido, b.Valor_Liquido_Final "
            "FROM anticipation_db.dbo.anticipation_db AS a "
            "INNER JOIN anticipation_db.dbo.borderos AS b "
            "    ON a.Bordero = b.Bordero "
            "WHERE "
            "    a.Is_inserted = 1 "
            "    AND a.Control_id = 4 "
            "    AND a.created_at >= '2026-04-17' "
            "    AND a.created_at < DATEADD(DAY, 1, '2026-04-17')"
        )
        colunas = [col[0] for col in cursor.description]
        linhas = cursor.fetchall()
        return [dict(zip(colunas, row)) for row in linhas]

    def buscar_dados_para_rpa(self, cedentes: list[str]) -> pd.DataFrame:
        if not self._conn:
            raise ConnectionError("Conexão não estabelecida. Chame conectar() primeiro.")

        placeholders = ",".join("?" for _ in cedentes)

        query = (
            "SELECT "
            "    a.Bordero, a.Cedente, a.Sacado, a.Titulo, a.Valor, "
            "    a.Emissao, a.created_at, a.Vencimento, "
            "    b.Valor_Liquido, b.Valor_Liquido_Final, "
            "    b.Valor_Total_Desagio, b.Debito_Credito "
            "FROM anticipation_db.dbo.anticipation_db AS a "
            "INNER JOIN anticipation_db.dbo.borderos AS b "
            "    ON a.Bordero = b.Bordero "
            "WHERE "
            "    a.Is_inserted = 1 "
            "    AND a.Control_id = 4 "
            "    AND a.created_at >= '2026-04-17 00:00:00' "
            "    AND a.created_at < '2026-04-18 00:00:00' "
            f"    AND a.Cedente IN ({placeholders}) "
            "    AND b.Valor_Liquido_Final IS NOT NULL "
            "    AND b.Valor_Total_Desagio IS NOT NULL"
        )

        cursor = self._conn.cursor()
        cursor.execute(query, cedentes)

        colunas = [col[0] for col in cursor.description]
        linhas = cursor.fetchall()

        return pd.DataFrame([dict(zip(colunas, row)) for row in linhas])

    def atualizar_valor_liquido(self, numero_bordero: int, valor: float):
        if not self._conn:
            raise ConnectionError("Conexão não estabelecida. Chame conectar() primeiro.")

        cursor = self._conn.cursor()
        cursor.execute(
            "UPDATE [anticipation_db].[dbo].[borderos] "
            "SET Valor_Liquido_Final = ? "
            "WHERE Bordero = ?",
            valor,
            numero_bordero,
        )
        self._conn.commit()
        print(f"[DB] Bordero {numero_bordero} atualizado com valor R$ {valor:.2f}")

    def atualizar_desagio_e_debito_credito(
        self,
        numero_bordero: int,
        valor_total_desagio: float,
        debito_credito: float,
    ) -> None:
        if not self._conn:
            raise ConnectionError("Conexão não estabelecida. Chame conectar() primeiro.")

        cursor = self._conn.cursor()
        cursor.execute(
            "UPDATE [anticipation_db].[dbo].[borderos] "
            "SET Valor_Total_Desagio = ?, Debito_Credito = ? "
            "WHERE Bordero = ?",
            valor_total_desagio,
            debito_credito,
            numero_bordero,
        )
        self._conn.commit()
        print(
            f"[DB] Bordero {numero_bordero}: Valor_Total_Desagio={valor_total_desagio:.2f} "
            f"Debito_Credito={debito_credito:.2f}"
        )

    def atualizar_debito_credito(self, numero_bordero: int, debito_credito: float) -> None:
        """Atualiza só ``Debito_Credito``; ``Valor_Total_Desagio`` permanece o cadastrado no banco."""
        if not self._conn:
            raise ConnectionError("Conexão não estabelecida. Chame conectar() primeiro.")

        cursor = self._conn.cursor()
        cursor.execute(
            "UPDATE [anticipation_db].[dbo].[borderos] "
            "SET Debito_Credito = ? "
            "WHERE Bordero = ?",
            debito_credito,
            numero_bordero,
        )
        self._conn.commit()
        print(f"[DB] Bordero {numero_bordero}: Debito_Credito={debito_credito:.2f}")


db_service = DatabaseService()