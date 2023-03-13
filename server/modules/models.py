import sqlalchemy
from sqlalchemy.orm import declarative_base


Base = declarative_base()


class Test(Base):
    __tablename__ = 'test'

    id = sqlalchemy.Column(sqlalchemy.Integer, primary_key=True)
    name = sqlalchemy.Column(sqlalchemy.String(length=40))


def create_tables(engine):
    Base.metadata.create_all(engine)
