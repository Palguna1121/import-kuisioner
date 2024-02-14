import pandas as pd
import pymysql

def connect_to_mysql():
    return pymysql.connect(host='localhost',
                           user='root',
                           password='',
                           database='kuisioner')

def clean_excel_data(data):
    cleaned_data = data.fillna('')
    return cleaned_data

def insert_audien_answers_to_mysql(data):
    connection = connect_to_mysql()
    cursor = connection.cursor()
    
    for index, row in data.iterrows():
        cursor.execute("INSERT INTO audien_answers (email, nama, usia, jenis_kelamin, tingkat_pendidikan, masa_kerja) VALUES (%s, %s, %s, %s, %s, %s)",
                       (row['Email'], row['Nama'], row['Usia'], row['Jenis Kelamin'], row['Tingkat Pendidikan'], row['Masa Kerja']))
    
    connection.commit()
    connection.close()

def insert_responses_to_mysql(data):
    connection = connect_to_mysql()
    cursor = connection.cursor()
    
    for index, row in data.iterrows():
        email = row['Email']
        for column in data.columns[6:]:
            question_id = data.columns.get_loc(column) - 5
            
            answer_text = row[column]
            base_answer_id = (question_id - 1) * 5
            if answer_text == 'STS = Sangat Tidak Setuju':
                answer_id = base_answer_id + 1
            elif answer_text == 'TS = Tidak Setuju':
                answer_id = base_answer_id + 2
            elif answer_text == 'N = Netral':
                answer_id = base_answer_id + 3
            elif answer_text == 'S = Setuju':
                answer_id = base_answer_id + 4
            elif answer_text == 'SS = Sangat Setuju':
                answer_id = base_answer_id + 5
            
            cursor.execute("INSERT INTO responses (email, question_id, answer_id) VALUES (%s, %s, %s)", (email, question_id, answer_id))
    
    connection.commit()
    connection.close()

def read_excel(filename):
    return pd.read_excel(filename)

if __name__ == "__main__":
    excel_filename = 'responses50.xlsx'
    excel_data = read_excel(excel_filename)
    
    cleaned_data = clean_excel_data(excel_data)
    
    insert_audien_answers_to_mysql(cleaned_data[['Email', 'Nama', 'Usia', 'Jenis Kelamin', 'Tingkat Pendidikan', 'Masa Kerja']])
    
    insert_responses_to_mysql(cleaned_data)
