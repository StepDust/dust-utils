class MySQLClient:
    def __init__(self, host, port, user, password, database, charset="utf8mb4"):
        # 只有在实例化 MySQLClient() 时才会执行 import
        try:
            import pymysql

            self.pymysql = pymysql
        except ImportError:
            raise ImportError(
                "检测到未安装 pymysql。请执行 'pip install pymysql' 以使用此功能。"
            )

        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.database = database
        self.charset = charset
        self.connection = None

    def connect(self):
        if not self.connection:
            self.connection = self.pymysql.connect(
                host=self.host,
                port=self.port,
                user=self.user,
                password=self.password,
                database=self.database,
                charset=self.charset,
            )

    def query(self, sql):
        self.connect()
        with self.connection.cursor(self.pymysql.cursors.DictCursor) as cursor:
            cursor.execute(sql)
            return cursor.fetchall()

    def execute(self, sql):
        self.connect()
        with self.connection.cursor() as cursor:
            cursor.execute(sql)
            self.connection.commit()
            return cursor.rowcount

    def close(self):
        if self.connection:
            self.connection.close()
            self.connection = None


# 示例
if __name__ == "__main__":
    db = MySQLClient(
        host="localhost", port=3306, user="root", password="123456", database="test"
    )

    # 查询（使用 f-string 拼接）
    age = 18
    sql = f"SELECT * FROM users WHERE age > {age}"
    data = db.query(sql)
    print(data)

    # 插入
    name = "Alice"
    age = 22
    insert_sql = f"INSERT INTO users(name, age) VALUES('{name}', {age})"
    affected = db.execute(insert_sql)
    print("影响行数:", affected)

    db.close()
