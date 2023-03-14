import json
import os
import mysql.connector
from mysql.connector import errorcode
import numpy
from datetime import datetime
import re

DATABASE_CONFIG_PATH = './config/database_config.json'

def check_empty(text):
    if len(text)!=0:
        return True
    else:
        return False
# def create_folder():
#     if os.path.exists('Output'):
#         None
#     else:
#         folder_path = 'Output'
#         os.mkdir(folder_path)
def check_phone(string_phone):
    try:
        phone = int(string_phone)
        if re.search('^[0-9]{10}$', string_phone):
            return True
        else:
            return False
    except:
        return False


def check_email(string_email):
    try:
        if re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b', string_email):
            return True
        else:
            return False
    except:
        return False


def split_date_time():
    date_time = str(datetime.now())
    date_split = date_time.split(' ')
    (h, mi, s) = date_split[1].split(':')
    (y, m, d) = date_split[0].split('-')
    return d, m, y, h, mi, s


def divide_chunks(list_data, number):
    for i in range(0, len(list_data), number):
        yield list_data[i:i + number]



def check_exit_file(file_path):
    if not os.path.exists(file_path):
        print(file_path + " is not exit")
        return False
    else:
        return True


def load_config(config_path):
    data = None

    if check_exit_file(config_path):
        with open(config_path, 'r', encoding='cp932', errors='ignore') as config_file:
            data = json.load(config_file)
        return data
    else:
        return False


def open_note_pad(path):
    cmd_string = 'notepad.exe ' + path
    os.system(cmd_string)


data_config = load_config(DATABASE_CONFIG_PATH)
try:
    cnx = mysql.connector.connect(
        host=data_config['host'],
        user=data_config['user'],
        password=data_config['password'],
        port=data_config['port'],
        database=data_config['database']
    )
    cursor = cnx.cursor()
except mysql.connector.Error as err:
    if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
        print("something is wrong with username or password")
    elif err.errno == errorcode.ER_BAD_DB_ERROR:
        print("data does not exist")
    else:
        print(err)
else:
    print("Connection database successfully")

def select_data(table):
    try:
        query = f"select *from {data_config['database']}.{table}"
        cursor.execute(query)
        data = cursor.fetchall()
        result = numpy.array(data)
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
        exit(1)
    return result


def search_data(name, table):
    try:
        query = f"select *from {data_config['database']}.{table} as h where h.name like '%{name}%'"
        cursor.execute(query)
        data = cursor.fetchall()
        result = numpy.array(data)
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
        exit(1)
    return result


def insert_data(name, phone, address, description):
    query = f"INSERT INTO {data_config['database']}.hospital (name, phone, address, description) VALUES(%s,%s,%s,%s)"
    args = (name, phone, address, description)
    try:
        cursor.execute(query, args)
        cnx.commit()
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
        return False
    else:
        return True


def delete_data(id, table):
    query = f"delete from {data_config['database']}.{table} where id ={id}"
    try:
        cursor.execute(query)
        cnx.commit()
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
        exit(1)
        return False
    else:
        return True


def select_doctor_and_name_hospital():
    query = "select do.id ,do.name ,do.phone ,do.email ,do.address ,ho.name " \
            f"from {data_config['database']}.hospital as ho " \
            f"inner join {data_config['database']}.doctor as do " \
            "on ho.id=do.hospital_id  " \
            "order by do.id ASC"
    try:
        cursor.execute(query)
        data = cursor.fetchall()
        result = numpy.array(data)
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
    return result


def select_patient_and_name_hospital():
    try:
        query = "select pa.id ,pa.name ,pa.phone ,pa.email ,pa.address ,ho.name " \
                f"from {data_config['database']}.hospital as ho " \
                f"inner join {data_config['database']}.patient as pa " \
                "on ho.id=pa.hospital_id  " \
                "order by pa.id ASC"
        cursor.execute(query)
        data = cursor.fetchall()
        result = numpy.array(data)
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
    return result


def select_schedule():
    try:
        query = "select sc.id ,sc.name ,sc.`date` ,do.name ,pa.name ,sc.`result`  " \
                f"from {data_config['database']}.doctor as do  " \
                f"inner join {data_config['database']}.schedule  as sc " \
                "on do.id =sc.doctor_id  " \
                f"inner join {data_config['database']}.patient as pa " \
                "on pa.id =sc.patient_id " \
                "order by sc.id ASC"

        cursor.execute(query)
        data = cursor.fetchall()
        result = numpy.array(data)
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
    return result


def insert_doctor(name, phone, email, address, hospital_id, table):
    query = f"INSERT INTO {data_config['database']}.{table} (name, phone, email, address, hospital_id) VALUES(%s,%s,%s,%s,%s);"
    args = (name, phone, email, address, hospital_id)
    try:
        cursor.execute(query, args)
        cnx.commit()
        return True
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
        return False


def edit_data(id, name, phone, address, description):
    query = f"UPDATE {data_config['database']}.hospital " \
            f"SET name='{name}', phone='{phone}', address='{address}', description='{description}'" \
            f" WHERE id={id};"
    try:
        cursor.execute(query)
        cnx.commit()
        return True
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
        return False


def edit_data_doctor(id, name, phone, email, address, hospital_id, table):
    query = f"UPDATE {data_config['database']}.{table} " \
            f"SET name='{name}', phone='{phone}', email='{email}', address='{address}', hospital_id={hospital_id}" \
            f" WHERE id={id};"
    try:
        cursor.execute(query)
        cnx.commit()
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
        return False
    else:
        return True


def edit_data_schedule(id, name, date, doctor_id, patient_id, resulut):
    query = f"UPDATE {data_config['database']}.schedule " \
            f"SET name='{name}', `date`='{date}', doctor_id={doctor_id}, patient_id={patient_id}, `result`='{resulut}' " \
            f"WHERE id={id};"
    try:
        cursor.execute(query)
        cnx.commit()
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
        return False
    else:
        return True


def insert_schedule(name, date, doctor_id, patient_id):
    query = f"INSERT INTO {data_config['database']}.schedule (name, `date`, doctor_id, patient_id) VALUES(%s,%s,%s,%s);"
    args = (name, date, doctor_id, patient_id)
    try:
        cursor.execute(query, args)
        cnx.commit()
        return True
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
        return False


def search_schedule(date_from, date_to, name):
    query = "SELECT  ne.id, ne.name, ne.date ,do.name, pa.name , ne.result " \
            f"FROM {data_config['database']}.doctor AS do " \
            f"INNER JOIN (SELECT * FROM {data_config['database']}.schedule AS sc " \
            f"WHERE sc.`date`  BETWEEN '{date_from}'AND '{date_to}' " \
            f"AND sc.name LIKE '%{name}%') AS ne  " \
            "ON ne.doctor_id=do.id " \
            f"INNER JOIN {data_config['database']}.patient AS pa " \
            "ON pa.id =ne.patient_id"
    try:
        cursor.execute(query)
        data = cursor.fetchall()
        result = numpy.array(data)
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
    return result



def convert_name_to_id(name, table):
    query = f"select h.id from {data_config['database']}.{table} h where h.name='{name}'"
    try:
        cursor.execute(query)
        data = cursor.fetchall()
        result = numpy.array(data)

    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
    return result


def select_table_name(table):
    query = f"select {table}.name from {data_config['database']}.{table}"
    try:
        cursor.execute(query)
        data = cursor.fetchall()
        result = numpy.array(data)
    except mysql.connector.Error as err1:
        print("Fail: {}".format(err1))
    return result
