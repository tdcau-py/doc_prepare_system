import sqlalchemy
from sqlalchemy.orm import sessionmaker

from models import create_tables


def create_database(db_user: str, db_user_password: str, db_host: str, db_port: str, db_name:str):
    """Создание базы данных"""
    DSN = f'mysql://{db_user}:{db_user_password}@{db_host}:{db_port}'
    engine = sqlalchemy.create_engine(DSN)
    conn = engine.connect()
    conn.execute('COMMIT;')
    conn.execute(f'CREATE DATABASE IF NOT EXISTS {db_name};')

    return conn.close()


def db_connection(db_user: str, db_user_password: str, db_host: str, db_port: str, db_name:str):
    """Подключение к БД"""
    DSN = f'mysql://{db_user}:{db_user_password}@{db_host}:{db_port}/{db_name}'
    engine = sqlalchemy.create_engine(DSN)

    return engine


# create_database()

# create_tables(engine=engine)

# Session = sessionmaker(bind=engine)
# session = Session()

# session.close()
