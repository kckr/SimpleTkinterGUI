import pyodbc


class DbUtil:
    dbname = None
    tablename1 = None
    tablename2 = None  # table names
    conn = None

    def __init__(self, dbname, tablename1, tablename2):  # class constructor
        self.dbname = dbname  #
        self.tablename1 = tablename1
        self.tablename2 = tablename2
        self.conn = None

    def connect(self):  # initiates the connection
        # connection string
        conn_str = """Driver={ODBC Driver 13 for SQL Server}; \
                                             Server=*****;\
                                             Database=""" + self.dbname + """; \
                                             UID=****; \
                                             PWD=****; \
                                            """
        self.conn = pyodbc.connect(conn_str, autocommit=True)  # pyodbc connection
        cursor = self.conn.cursor()  # cursor object
        return cursor

    def readdata(self, pin):  # read function by accepting pin
        cursor = self.connect()
        # pin = 'BF80A0206A01000000'
        # sets record from SQL Query

        record = cursor.execute("""SELECT  [ASSEMBLY_ITEM_NO]
                                ,[COMPONENT_NO]
                                ,[COMPONENT_DESC]                               
                                ,[QUANTITY]   
                                FROM [""" + self.dbname + """].[dbo].[""" + self.tablename1 + """] 
                                where [ASSEMBLY_ITEM_NAME] = '""" + pin + """'""").fetchall()

        cursor.close()  # closing the cursor
        self.conn.close()  # closing the connection
        print(record)

        return record

    def readTable2(self, pin):
        cursor = self.connect()
        record = cursor.execute("""SELECT [ASSEMBLY_ITEM_DESC]
                                FROM [""" + self.dbname + """].[dbo].[""" + self.tablename2 + """]
                                where [ASSEMBLY_ITEM_NO] = '""" + pin + """'""").fetchall()
        cursor.close()  # closing the cursor
        self.conn.close()  # closing the connection
        return record

    def read_component(self, pin):
        cursor = self.connect()
        record = cursor.execute("""SELECT [ASSEMBLY_ITEM_NO]
                                ,[COMPONENT_DESC]
                                FROM [""" + self.dbname + """].[dbo].[""" + self.tablename1 + """]
                                where [COMPONENT_NO] = '""" + pin + """'""").fetchall()
        cursor.close()  # closing the cursor
        self.conn.close()  # closing the connection
        return record

    def read_description(self, pin):
        cursor = self.connect()
        record = cursor.execute("""SELECT 
                                [ASSEMBLY_ITEM_DESC]
                                FROM [""" + self.dbname + """].[dbo].[""" + self.tablename2 + """]
                                where [ASSEMBLY_ITEM_NO] = '""" + pin + """'""").fetchall()
        cursor.close()  # closing the cursor
        self.conn.close()  # closing the connection
        return record

    def assemblyitemsno(self):
        cursor = self.connect()
        record = cursor.execute("""SELECT 
                                   [ASSEMBLY_ITEM_NO]
                                   FROM [XML].[dbo].[ASSEMBLY_ITEM]""").fetchall()
        cursor.close()  # closing the cursor
        self.conn.close()  # closing the connection
        return record

# mydb = DbUtil('XML', 'COMPONENT_ROW', 'ASSEMBLY_ITEM')
# myrecord = mydb.readdata("BF80A0206A01000000")
# print(myrecord)
#
