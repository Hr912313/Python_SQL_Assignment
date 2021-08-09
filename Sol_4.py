#installing the libraries psycopg2, Workbook, pandas
import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd

class employees:

    def emp(self):
        try:
            # to connect to the PostgreSQL database server in the Python program using the psycopg database adapter.
            conn = psycopg2.connect(
                database="python-SQL",
                user="hr3562",
                password="harshraj")
             
            # Creating a cursor object using the cursor() method
            cursor = conn.cursor()
            # Reading table which we imported using connection through query
            query = """
                    select dept.deptno, dept_name, sum(total_compensation) from Compensation, dept
                    where Compensation.dept_name=dept.dname
                    group by dept_name, dept.deptno
                    """

            cursor.execute(query)
            #iterating inside description
            columns = [desc[0] for desc in cursor.description]
            data = cursor.fetchall()
            df = pd.DataFrame(list(data), columns=columns)
            #storing values inside excel
            writer = pd.ExcelWriter('ques_4.xlsx')
            # converting data frame to excel
            df.to_excel(writer, sheet_name='bar')
            writer.save()

        except Exception as e:
            print("Error has occurred", e)

        finally:

            if conn is not None:
                cursor.close()
                conn.close()


if __name__ == '__main__':
    conn = None
    cursor = None
    employee = employees()
    employee.emp()
