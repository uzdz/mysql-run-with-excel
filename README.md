# mysql-run-with-excel

通过按照Excel约定执行SQL语句的小工具。

![avatar](images/d1.png)

注意事项如下：

1. 数据库连接地址、数据库端口号、用户名、密码、表名、执行SQL必须存在，且位于第一和第三行，位置不能改动。
2. 执行多条SQL语句通过`|`进行拼接字符串。
3. sheet页名称为数据库名称。