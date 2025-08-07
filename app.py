from flask import Flask, render_template, request, redirect
import mysql.connector

app = Flask(__name__)

# MySQL DB configuration
db_config = {
    'host': 'localhost',
    'user': 'root',       # your MySQL username
    'password': '1234',   # your MySQL password
    'database': 'school_sms'
}

# Home route: registration form
@app.route('/', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        admission = request.form['admission']
        assesment = request.form['assesment']
        fname = request.form['fname']
        lname = request.form['lname']
        dob = request.form['dob']
        gender = request.form['gender']
        class_name = request.form['class']
        parent = request.form['parent']
        contact = request.form['contact']

        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor()
        sql = """
        INSERT INTO learners 
        (admission_number,assesment_number, first_name, last_name, date_of_birth, gender, class_name, parent_name, parent_contact)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        values = (admission,assesment, fname, lname, dob, gender, class_name, parent, contact)
        try:
            cursor.execute(sql, values)
            conn.commit()
        except mysql.connector.IntegrityError:
            return "Error: Duplicate Admission Number"
        conn.close()
        return redirect('/learners')
    return render_template('register.html')

# View learners
@app.route('/learners')
def learners():
    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT * FROM learners ORDER BY id DESC")
    data = cursor.fetchall()
    conn.close()
    return render_template('learners.html', learners=data)

if __name__ == '__main__':
    app.run(debug=True)
