#installing the libraries psycopg2, Workbook,pandas
import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd

class employees:
    def emp(self):
        # to connect to the PostgreSQL database server in the Python program using the psycopg database adapter.
        try:
            conn = psycopg2.connect(
                host="localhost",
                database="python-SQL",
                user="hr3562",
                password="harshraj")
            # Creating a cursor object using the cursor() method
            cursor = conn.cursor()
            # Reading table which we imported using connection through query
            query= """SELECT e1.empno, e1.ename, (case when mgr is not null then (select ename from emp as e2 where e1.mgr=e2.empno limit 1) else null end) as manager
            from emp as e1"""
            cursor.execute(query)

            columns = [desc[0] for desc in cursor.description]
            data = cursor.fetchall()
            df = pd.DataFrame(list(data), columns=columns)
            # storing values inside excel
            writer = pd.ExcelWriter('ques_1.xlsx')
            # converting data frame to excel
            df.to_excel(writer, sheet_name='bar')
            writer.save()

        except Exception as e:
            print("Error is present", e)
        finally:

            if conn is not None:
                cursor.close()
                conn.close()


if __name__=='__main__':
    conn = None
    cursor = None
    employee = employees()
    employee.emp()
