import csv
from docxtpl import DocxTemplate
import psycopg2

conn = psycopg2.connect(
        dbname='reviews',
        user='postgres',
        password='postgres',
        host='127.0.0.1',
        port='5432',
    )
cursor = conn.cursor()
cursor.execute('CREATE TABLE IF NOT EXISTS reviews (id SERIAL PRIMARY key, fio TEXT, work_name TEXT, review TEXT)')

r_file = open("data.csv", encoding='utf-8')
file_reader = csv.reader(r_file, delimiter=",")
next(file_reader)

for row in file_reader:

    doc = DocxTemplate('my_template.docx')

    context = {
        'uchebnaya_kafedra': row[1],
        'the_name_of_the_work': row[2],
        'kurs_number': row[3],
        'group_number': row[4],
        'forma_obycheniya': row[5],
        'fio_stud': row[6],
    }

    doc.render(context)

    doc.save('Отчет о работе студента %s.docx' % (row[6]))

    review = input('Введите отзыв на курсовую работу "%s" студента %s:' % (row[2], row[6]))

    cursor.execute("INSERT INTO reviews (fio, work_name, review) VALUES ('%s', '%s', '%s')" % (row[6], row[2], review))
    conn.commit()

conn.close()
