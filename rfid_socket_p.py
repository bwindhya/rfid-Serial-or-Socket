import socket
import time
import datetime
import openpyxl
import psycopg2

wb = openpyxl.load_workbook(r"D:\RFID\mapping_unit.xlsx")
ws = wb.active

ip = '10.30.246.2'
port = 6000

conn = psycopg2.connect(database="db_vhms", user="vhms_db_user", password="macan123", host="10.30.241.4", port="5432")
print('Opened database succesfully')
cur = conn.cursor()

def connect_reader():
    try:
        reader_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        reader_socket.connect((ip, port))
        print("Connected to RFID reader.")
        return reader_socket
    except Exception as e:
        print(f"Error Connection to RFID: {e}")
        return reader_socket

def read_rfid(reader_socket):
    try:
        reader_socket.send(b'READ')  # Send a command to the reader to read a tag
        tag_data = reader_socket.recv(1024)  # Receive tag data from the reader
        return tag_data
    except Exception as e:
        print(f"Error read RFID: {e}")
        return None

def save_rfid(kode, no_unit, time):
    sql = (""" INSERT INTO rfid (kode_tag, serial_unit, calendar) VALUES ('{0}', '{1}', '{2}');""".format(kode, no_unit, time))
    cur.execute(sql)
    conn.commit()

def save_cycle(kode, no_unit, tanggal, waktu):
    sql = f"insert into cycle_rfid (kode_tag, serial_unit, tanggal, waktu) values ('{kode}', '{no_unit}', '{tanggal}', '{waktu}')"
    cur.execute(sql)
    conn.commit()

def eksekusi(index):
    hasil = data[index:index+6]
    for row in reader:
        if hasil == row[3]:
            ambil_calen = f"select calendar from rfid where kode_tag = '{hasil}' and calendar between '{old_time}' and '{current_time}'"
            cur.execute(ambil_calen)
            result = [r[0] for r in cur.fetchall()]
            if result == []:
                print(row[3], row[0], current_time)
                save_rfid(row[3], row[0], current_time)
                save_cycle(row[3], row[0], current_tgl, current_waktu)
                time.sleep(2)
            elif (old_time_convert <= min(result)):
                print(f'{hasil} can not be saved, wait another {detik-now.second} seconds')      
                time.sleep(2)             

reader_socket = connect_reader()
if reader_socket:
    try:
        n = 0
        while True:
            n += 1
            data = read_rfid(reader_socket).hex()
            if data:
                detik = 120
                now = datetime.datetime.now()
                old = now - datetime.timedelta(seconds=detik)
                current_time = now.strftime('%Y-%m-%d %H:%M:%S')
                current_tgl = now.date()
                current_waktu = now.strftime('%H:%M:%S')
                old_time = old.strftime('%Y-%m-%d %H:%M:%S')
                old_time_convert = datetime.datetime.strptime(old_time, '%Y-%m-%d %H:%M:%S')

                reader = ws.iter_rows(min_row=2, values_only=True)
                if '0013' in data:
                    index_awal = data.index('0013')
                    eksekusi(index_awal)
                elif '0012' in data:
                    index_awal2 = data.index('0012')
                    eksekusi(index_awal2)
                elif '0016' in data:
                    index_awal3 = data.index('0016')
                    eksekusi(index_awal3)
                elif '0017' in data:
                    index_awal4 = data.index('0017')
                    eksekusi(index_awal4)
                elif '0018' in data:
                    index_awal5 = data.index('0018')
                    eksekusi(index_awal5)
                elif '0019' in data:
                    index_awal7 = data.index('0019')
                    eksekusi(index_awal7)
                elif '0011' in data:
                    index_awal8 = data.index('0011')
                    eksekusi(index_awal8)
                elif '0015' in data:
                    index_awal15 = data.index('0015')
                    eksekusi(index_awal15)
            else:
                print(f'{n}. Idling...')
                time.sleep(2)
    except KeyboardInterrupt:
        print("Exiting...")
    finally:
        reader_socket.close()
