import pandas as pd
import sqlalchemy as al
from sqlalchemy.exc import SQLAlchemyError
from xlrd.biffh import XLRDError


# Excel data recording class in MySql
class ConvertFile:
    def __init__(self, base, user, host, table, patch_to_excel, work_shit, password=''):
        self.base = base
        self.user = user
        self.host = host
        self.table = table
        self.patch_to_excel = patch_to_excel
        self.work_shit = work_shit
        self.password = password
        self.data_excel = None

    def edit_table(self, replace=True):
        try:
            # reading the data from the letter Excel
            self.data_excel = pd.read_excel(self.patch_to_excel, sheet_name=self.work_shit)
        except XLRDError:
            return False
        else:
            if replace is True:
                value = 'replace'
            else:
                value = 'append'
            try:
                try:
                    # creating connection with MySql database
                    engine = al.create_engine(f'mysql+pymysql://{self.user}:{self.password}@{self.host}/{self.base}')

                    # writing data in MySql
                    with engine.connect() as con, con.begin():
                        self.data_excel.to_sql(self.table, con, if_exists=value)
                        self.data_excel = None
                        return True
                except UnicodeEncodeError:
                    return False
            except SQLAlchemyError:
                return False


# if __name__ == '__main__':
#     d = ConvertFile('Tasks', 'Volodimir', 'localhost', 'Tr', 'D:\\Робота\\AMAZON EMP_2.xls', 'Products', '10091984')
#     res = d.edit_table()
#     if res is False:
#         print('f')
#     else:
#         print('t')
