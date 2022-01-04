from flask import Flask, render_template, url_for, request, redirect
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
from sqlalchemy import or_
from datetime import datetime
from pathlib import Path
import os
import xlrd
import xlwt
from werkzeug.utils import redirect

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///parts.db'
db = SQLAlchemy(app)

class Teile(db.Model):
    id = db.Column(db.Integer, primary_key = True)
    name = db.Column(db.String(50), nullable = False)
    description = db.Column(db.String(200))
    place = db.Column(db.String(10), nullable = False)
    number = db.Column(db.Integer, nullable = False)

    def __repr__(self):
        return '<Teile %r>' % self.id

@app.route('/', methods=['POST', 'GET'])
def index():
    if request.method == 'POST':
        if request.form['action']=='Suche':
            search_input = request.form['search']
            search = "%{}%".format(search_input)
            parts = Teile.query.filter(or_(Teile.name.like(search), Teile.description.like(search))).all()
            
            if Teile.query.filter(Teile.name.like(search)).first()== None and Teile.query.filter(Teile.description.like(search)).first()== None:
                return render_template('nothing_found.html')
            else: 
                return render_template('search_results.html', parts = parts)
            
        
        elif request.form['action']=='Hinzufügen':
            return render_template('add.html')

        elif request.form['action']=='Datenbank importieren':
            return render_template('upload.html')

        elif request.form['action']=='Datenbank exportieren':
            
            downloads_path = str(Path.home())+"/Downloads/"

            wb = xlwt.Workbook()
            ws = wb.add_sheet('Database')

            number_of_entries = Teile.query.count()

            for i in range(number_of_entries):
                part = Teile.query.get(i+1)

                ws.write(i+1,0,i+1)
                ws.write(i+1,1,part.name)
                ws.write(i+1,2,part.description)
                ws.write(i+1,3,part.place)
                ws.write(i+1,4,part.number)

            wb.save(downloads_path+"Teile_Exportiert.xls")

            return redirect('/')
        else: 
            return "something went wrong"
        
            
    else: 
        parts = Teile.query.order_by(Teile.id).all()
        return render_template('index.html', parts = parts)

@app.route('/add', methods = ['POST', 'GET'])
def add():
    if request.method == 'POST':
        part_name = request.form['name']
        part_description = request.form['description']
        part_place = request.form['place']
        part_number = request.form['number']

        new_part = Teile(
            name = part_name, 
            description = part_description,
            place = part_place,
            number = part_number            
            )

        db.session.add(new_part)
        db.session.commit()

        return redirect('/')


@app.route('/delete/<int:id>')
def delete(id):
    part_to_delete = Teile.query.get_or_404(id)

    try: 
        db.session.delete(part_to_delete)
        db.session.commit()
        return redirect('/')
    except:
        return 'There was a problem deleting that part'

@app.route('/update/<int:id>', methods=['POST', 'GET'])
def update(id):
    part = Teile.query.get_or_404(id)
    if request.method == 'POST':
        name = request.form['name']
        description = request.form['description']
        place = request.form['place']
        number = request.form['number']
        if name != '':
            part.name = name
        if description != '':
            part.description = description
        if place != '':
            part.place = place
        if number != '':
            part.number = number

        try: 
            db.session.commit()
            return redirect('/')
        except: 
            return "there was an issue updating ur task"
    else: 
        return render_template('update.html', part = part)

@app.route('/upload', methods = ['POST'])
def upload():

    parent_path = Path(__file__).parent.resolve()

    file = request.files['datei']
    file.save(str(parent_path)+"/Teile.xls")

    db.session.query(Teile).delete()
    db.session.commit()

    loc = str(parent_path)+"/Teile.xls"
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    
    for i in range(sheet.nrows-1):
        part_name = sheet.cell_value(i+1,1)
        part_description = sheet.cell_value(i+1,2)
        part_place = sheet.cell_value(i+1,3)
        part_number = sheet.cell_value(i+1,4)

        new_part = Teile(
            name = part_name, 
            description = part_description,
            place = part_place,
            number = part_number            
            )

        db.session.add(new_part)
        db.session.commit()
    
    os.remove(str(parent_path)+"/Teile.xls")

    return redirect('/')

if __name__ == "__main__":
    app.run(host= "localhost", port = 3000, debug=True)