import sqlite3


def create_user_table(con, cur):
    statement = '''
    CREATE TABLE IF NOT EXISTS user (
        user TEXT PRIMARY KEY,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    '''
    cur.execute(statement)
    con.commit()


def create_interaction_table(con, cur):
    statement = '''
    CREATE TABLE IF NOT EXISTS interaction (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user TEXT NOT NULL,
        account_number TEXT NOT NULL,
        statement_month TEXT NOT NULL,
        status TEXT NOT NULL,
        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user) REFERENCES user(user)
    )'''
    cur.execute(statement)
    con.commit()


def main():
    connection = sqlite3.connect('statement_retrieval.db')
    cursor = connection.cursor()

    create_user_table(connection, cursor)
    create_interaction_table(connection, cursor)


if __name__ == '__main__':
    main()
    print('Database and table setup complete.')