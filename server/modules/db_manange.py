import sqlalchemy
from sqlalchemy.orm import sessionmaker

from models import create_tables

DSN = 'mysql://root:flvby@localhost:3306/test'
engine = sqlalchemy.create_engine(DSN)

create_tables(engine=engine)

Session = sessionmaker(bind=engine)
session = Session()

session.close()
