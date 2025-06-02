import sqlite3
import time
from typing import Any

class InteractionModel:
    def __init__(self):
        self.connection = None
        self.cursor = None

    def get_connection(self):
        conn = sqlite3.connect('statement_retrieval.db')
        conn.execute('PRAGMA journal_mode=WAL;')
        return conn
        

    def find_by_id(self, id : int):
        for _ in range(5):
            try:
                with self.get_connection() as conn:
                    cursor =  conn.cursor()
                    statement = 'SELECT * FROM interaction WHERE id = ?'
                    res = cursor.execute(statement, (id,))
                return res.fetchone()
            except sqlite3.OperationalError as e:
                if 'locked' in str(e):
                    time.sleep(1)
                else:
                    raise

    
    def find_interactions_by_user(self, user: str):
        for _ in range(5):
            try:
                with self.get_connection() as conn:
                    cursor = conn.cursor()
                    statement = 'SELECT * FROM interaction WHERE user = ?'
                    res = cursor.execute(statement, (user,))
                return res.fetchall()
            except sqlite3.OperationalError as e:
                if 'locked' in str(e):
                    time.sleep(1)
                else:
                    raise

    def find(self) -> list[Any]:
        for _ in range(5):
            try:
                with self.get_connection() as conn:
                    cursor = conn.cursor()
                    statement = 'SELECT * FROM interaction'
                    res = cursor.execute(statement)
                return res.fetchall()
            except sqlite3.OperationalError as e:
                if 'locked' in str(e):
                    time.sleep(1)
                else:
                    raise
    
    def insert(self, user: str, account_number : str, statement_month : str, status : str):
        for _ in range(5):
            try:
                with self.get_connection() as conn:
                    cursor = conn.cursor()
                    statement = 'INSERT INTO interaction (user, account_number, statement_month, status) VALUES (?, ?, ?, ?)'
                    cursor.execute(statement, (user, account_number, statement_month, status))
                    conn.commit()
                break
            except sqlite3.OperationalError as e:
                if 'locked' in str(e):
                    time.sleep(1)
                else:
                    raise


class UserModel:
    def __init__(self):
        self.connection = None
        self.cursor = None

    def get_connection(self):
        conn = sqlite3.connect('statement_retrieval.db')
        conn.execute('PRAGMA journal_mode=WAL;')
        return conn

    def find_by_id(self, id : str):
        for _ in range(5):
            try:
                with self.get_connection() as conn:
                    cursor = conn.cursor()
                    statement = 'SELECT * FROM user WHERE user = ?'
                    res = cursor.execute(statement, (id,))
                return res.fetchone()
            except sqlite3.OperationalError as e:
                if 'locked' in str(e):
                    time.sleep(1)
                else:
                    raise

    def find(self) -> list[Any]:
        for _ in range(5):
            try:
                with self.get_connection() as conn:
                    cursor = conn.cursor()
                    statement = 'SELECT * FROM user'
                    res = cursor.execute(statement)
                return res.fetchall()
            except sqlite3.OperationalError as e:
                if 'locked' in str(e):
                    time.sleep(1)
                else:
                    raise

    
    def insert(self, user : str):
        for _ in range(5):
            try: 
                with self.get_connection() as conn:
                    cursor = conn.cursor()
                    statement = 'INSERT INTO user (user) VALUES (?)'
                    cursor.execute(statement, (user,))
                    conn.commit()
                break
            except sqlite3.OperationalError as e:
                if 'locked' in str(e):
                    time.sleep(1)
                else:
                    raise