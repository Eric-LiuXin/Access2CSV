import pyodbc, sys, os, csv

def access2csv(file_path, shotname):
    csv_list = []

    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=%s;'%(file_path)
    )
    cnxn = pyodbc.connect(conn_str)
    crsr = cnxn.cursor()

    table_list = []
    for table_info in crsr.tables(tableType='TABLE'):
        table_list.append(table_info.table_name)

    for table in table_list:
        title_list = []
        for row in crsr.columns(table=table):
            title_list.append(row.column_name)

        sql = "SELECT %s FROM %s" %(', '.join(title_list), table)
        csv_path = r"%s-%s.csv" % (shotname, table)

        with open(csv_path, "w", newline="", encoding="utf-8") as csv_file:
            spam_writer = csv.writer(csv_file, dialect='excel')
            spam_writer.writerow(title_list)
            for row in crsr.execute(sql).fetchall():
                spam_writer.writerow(row)

        csv_list.append(csv_path)

    crsr.close()
    cnxn.close()

    return csv_list


def main():
    if len(sys.argv) < 2:
        print("请指定要转换的accdb文件。")
        return

    file_path = sys.argv[1]
    if not os.path.exists(sys.argv[1]):
        print("指定文件不存在，或不是有效路径。")
        return

    shotname, extension = os.path.splitext(file_path)
    if extension.lower() != '.accdb':
        print("请指定accdb类型的文件。")
        return

    csv_list = access2csv(file_path, shotname)
    print("CSV文件输出路径：%s" % ','.join(csv_list))

if __name__ == '__main__':
    main()