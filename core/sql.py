from sqlalchemy import create_engine,text
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.automap import automap_base
from contextlib import contextmanager
import threading
import pandas as pd
from core.config import settings
import random
class sql_orm():
    def __init__(self,DB_URL=settings.db_url):
        self.engine = create_engine(**DB_URL,query_cache_size=0)
        self.Base = automap_base()
        self.Base.prepare(self.engine, reflect=True)
        self.Session = sessionmaker(bind=self.engine)
    @contextmanager
    def session_scope(self):
        session = self.Session()
        try:
            yield session,self.Base
            session.commit()
        except Exception as e:
            session.rollback()
            print('有错误,回滚:'+str(e))
            raise
        finally:
            session.close()
    def add_data(self,df,table):
        with self.session_scope() as (sql, Base):
            pojo=getattr(Base.classes,table)
            rows = []
            for index, row in df.iterrows():
                temp = pojo(**row.to_dict())
                rows.append(temp)
            sql.bulk_save_objects(rows)
    def truncate_add_data(self, df, table):
        lock = threading.Lock()
        with lock:  # 使用线程锁
            with self.session_scope() as (sql, Base):
                pojo = getattr(Base.classes, table)
                try:
                    with sql.begin():  # 开始一个事务
                        sql.execute(text(f'TRUNCATE TABLE {table}'))
                        rows = []
                        for index, row in df.iterrows():
                            temp = pojo(**row.to_dict())
                            rows.append(temp)
                        sql.bulk_save_objects(rows)
                except Exception as e:
                    sql.rollback()  # 如果发生异常，回滚事务
                    raise e
    def get_cookies(self,id):
        with self.session_scope() as (sql, Base):
            pojo = getattr(Base.classes, "cookies")
            if id == "foura":
                suffix = random.choice([1, 2, 3])
                id = f"foura{suffix}"
            res = sql.query(pojo).filter(pojo.id==id).first()
            cookies_str = res.cookies
            cookies = {}
            try:
                for cookie in cookies_str.split(';'):
                    key, value = cookie.split('=', 1)
                    cookies[key] = value
            except:pass
            return {"cookies":cookies, "cookies_str":cookies_str}
    def get_data(self,table):
        df = pd.read_sql_table(table, con=self.engine)
        return df
    def get_engine(self):
        return self.engine
    def execute_sql(self,sql_query):
        with self.session_scope() as (sql, Base):
            sql.execute(text(sql_query))