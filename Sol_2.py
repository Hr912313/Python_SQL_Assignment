#installing the libraries psycopg2, Workbook,pandas
import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd

class Total_compensation:
    def compensation(self):
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
            query = """
            select emp.ename, emp.empno, dept.dname, (case when enddate is not null then ((enddate-startdate+1)/30)*(jobhist.sal) else ((current_date-startdate+1)/30)*(jobhist.sal) end)as Total_Compensation,
            (case when enddate is not null then ((enddate-startdate+1)/30) else ((current_date-startdate+1)/30) end)as Months_Spent from jobhist, dept, emp 
            where jobhist.deptno=dept.deptno and jobhist.empno=emp.empno"""
            cursor.execute(query)
            columns = [desc[0] for desc in cursor.description]
            data = cursor.fetchall()
            df = pd.DataFrame(list(data), columns=columns)
            # storing values inside excel
            writer = pd.ExcelWriter('ques_2.xlsx')
            #converting data frame to excel
            df.to_excel(writer, sheet_name='bar')
            writer.save()

        except Exception as e:
            print("Something went wrong", e)
        finally:

            if conn is not None:
                cursor.close()
                conn.close()


if __name__=='__main__':
    conn = None
    cursor = None
    comp = Total_compensation()
    comp.compensation()
