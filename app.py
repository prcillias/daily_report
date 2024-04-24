from flask import Flask, render_template, request, jsonify, send_file
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table
from datetime import datetime
import pandas as pd
import os
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import and_
from sqlalchemy import func

basedir = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] =\
        'sqlite:///' + os.path.join(basedir, 'database', 'database.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

class CustTemp(db.Model):
    id = db.Column(db.Integer, autoincrement=True, primary_key=True)
    date = db.Column(db.String(10))
    nama = db.Column(db.String(255))
    average_temp = db.Column(db.Float)
    safe_percentage = db.Column(db.Float)
    
    def __repr__(self):
        return f'<{self.nama}>'
    
class CustUnbalance(db.Model):
    id = db.Column(db.Integer, autoincrement=True, primary_key=True)
    date = db.Column(db.String(10))
    nama = db.Column(db.String(255))
    van = db.Column(db.Float)
    vbn = db.Column(db.Float)
    vcn = db.Column(db.Float)
    vab = db.Column(db.Float)
    vbc = db.Column(db.Float)
    vca = db.Column(db.Float)
    van_vbn = db.Column(db.Integer)
    van_vcn = db.Column(db.Integer)
    vbn_vcn = db.Column(db.Integer)
    vab_vbc = db.Column(db.Integer)
    vab_vca = db.Column(db.Integer)
    vbc_vca = db.Column(db.Integer)
    # van_vbn = db.Column(db.String(5))
    # van_vcn = db.Column(db.String(5))
    # vbn_vcn = db.Column(db.String(5))
    # vab_vbc = db.Column(db.String(5))
    # vab_ca = db.Column(db.String(5))
    # vbc_vca = db.Column(db.String(5))
    unbalance_count = db.Column(db.Integer)
    
    def __repr__(self):
        return f'<{self.nama}>'

@app.route('/')
def home():
    data = CustTemp.query.all()
    return render_template('index.html', data=data)

@app.route('/unbalance')
def unbalance():
    data = CustUnbalance.query.with_entities(CustUnbalance.nama).distinct().all()
    return render_template('unbalance.html', data=data)

@app.route('/get-data', methods=['GET'])
def get_data():
    date = request.args.get('date')

    data = CustTemp.query.filter(CustTemp.date == date).all()
    formatted_data = []
    if data:
        for item in data:
            formatted_item = {
                'nama': item.nama,
                'average_temp': item.average_temp,
                'safe_percentage': item.safe_percentage
            }
            formatted_data.append(formatted_item)
        return jsonify({"message": "success", "data": formatted_data})
    else:
        return jsonify({"message": "no data"})
    
@app.route('/get-data-voltage', methods=['GET'])
def get_data_voltage():
    date = request.args.get('date')
    cust = request.args.get('cust')

    data = CustUnbalance.query.filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).all()
    formatted_data = []
    if data:
        for item in data:
            formatted_item = {
                'nama': item.nama,
                'van': item.van,
                'vbn': item.vbn,
                'vcn': item.vcn,
                'vab': item.vab,
                'vbc': item.vbc,
                'vca': item.vca,
                'unbalance_count': item.unbalance_count,
            }
            formatted_data.append(formatted_item)
        return jsonify({"message": "success", "data": formatted_data})
    else:
        return jsonify({"message": "no data"})
    
@app.route('/clear-data', methods=['GET'])
def clear_data():
    date = request.args.get('date')
    data = CustTemp.query.filter(CustTemp.date == date).all()
    if data:
        for d in data:
            db.session.delete(d)
        db.session.commit()
        return jsonify({"message": "success"})
    else:
        return jsonify({"message": "no data"}) 
    
@app.route('/clear-data-voltage', methods=['GET'])
def clear_data_voltage():
    date = request.args.get('date')
    data = CustUnbalance.query.filter(CustUnbalance.date == date).all()
    if data:
        for d in data:
            db.session.delete(d)
        db.session.commit()
        return jsonify({"message": "success"})
    else:
        return jsonify({"message": "no data"}) 

@app.route('/upload', methods=['POST'])
def upload():
    nama = request.form.get('selectedCompany')
    date = request.form.get('date')
    uploaded_files = request.files.getlist('excel_files')
    data = []

    existing_data = CustTemp.query.filter_by(nama=nama, date=date).first()
    if existing_data:
        return jsonify({"message": "existed"})
    
    for index, excel_file in enumerate(uploaded_files):
        if excel_file.filename != '':
            df = pd.read_excel(excel_file)
            df.rename(columns={df.columns[-1]: "Status"}, inplace=True)
            data.append(df)
            
    if data:
        all_data = pd.concat(data, ignore_index=True)
        average_oil_temp = round(all_data['OilTemp'].mean(), 2)
        value_counts = all_data['Status'].value_counts()
        safe_percentage = round((value_counts.get("Safe", 0) / len(all_data)) * 100, 2)
        new = CustTemp(date=date, nama=nama, average_temp=average_oil_temp, safe_percentage=safe_percentage)
        db.session.add(new)
        db.session.commit()
        return jsonify({"message": "success", "date": date})
    else:
        return jsonify({"message": "failed"})
    
@app.route('/upload_voltage', methods=['POST'])
def upload_voltage():
    nama = request.form.get('selectedCompany')
    date = request.form.get('date')
    uploaded_files = request.files.getlist('excel_files')
    data = []

    existing_data = CustUnbalance.query.filter_by(nama=nama, date=date).first()
    if existing_data:
        return jsonify({"message": "existed"})
    
    
    for index, excel_file in enumerate(uploaded_files):
        if excel_file.filename != '':
            df = pd.read_excel(excel_file, usecols="H,I,Q,R,Z,AA")
            
            for _, row in df.iterrows():
                if abs(row['Van'] - row['Vbn']) > 15:
                    van_vbn = 1
                else:
                    van_vbn = 0
                if abs(row['Van'] - row['Vcn']) > 15:
                    van_vcn = 1
                else:
                    van_vcn = 0
                if abs(row['Vbn'] - row['Vcn']) > 15:
                    vbn_vcn = 1
                else:
                    vbn_vcn = 0
                if abs(row['Vab'] - row['Vbc']) > 15:
                    vab_vbc = 1
                else:
                    vab_vbc = 0
                if abs(row['Vab'] - row['Vca']) > 15:
                    vab_vca = 1
                else:
                    vab_vca = 0
                if abs(row['Vbc'] - row['Vca']) > 15:
                    vbc_vca = 1
                else:
                    vbc_vca = 0
                    
                # diff = abs(row['Van'] - row['Vbn']) + abs(row['Van'] - row['Vcn']) + abs(row['Vbn'] - row['Vcn'])
                unbalance_count = van_vbn + van_vcn + vbn_vcn + vab_vbc + vab_vca + vbc_vca
                
                new = CustUnbalance(date=date, nama=nama, van=row['Van'], vbn=row['Vbn'], vcn=row['Vcn'], vab=row['Vab'], vbc=row['Vbc'], vca=row['Vca'],
                                    van_vbn=van_vbn, van_vcn=van_vcn, vbn_vcn=vbn_vcn, vab_vbc=vab_vbc, vab_vca=vab_vca, vbc_vca=vbc_vca,
                                    unbalance_count=unbalance_count)
                db.session.add(new)
                db.session.commit()

    return jsonify({"message": "success", "date": date})
    
@app.route('/check', methods=['POST'])
def check():
    type = request.form.get('type')
    date = request.form.get('date')
    
    if type == 't':
        data = CustTemp.query.filter(CustTemp.date == date).all()
    elif type == 'v':
        cust = request.form.get('cust')
        data = CustUnbalance.query.filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).all()
        
    if data:
        return jsonify({'message': 'success'})    
    else:
        return jsonify({'message': 'failed'})    

@app.route('/make-pdf', methods=['POST'])
def make_pdf():
    date = request.form.get('date')
    date_obj = datetime.strptime(date, "%Y-%m-%d")

    formatted_date = date_obj.strftime("%d %B %Y")
    data = CustTemp.query.filter(CustTemp.date == date).all()

    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    p.setFont("Helvetica-Bold", 16)
    p.drawString(60, 765, f"Daily Report {formatted_date}")

    table_data = [['No', 'Nama', 'Suhu', 'Safe Percentage']]
    for i, row in enumerate(data):
        table_data.append([i+1, row.nama, str(row.average_temp) + '°C', str(row.safe_percentage) + '%'])

    col_widths = [35, 195, 122.5, 122.5]
    t = Table(table_data, colWidths=col_widths)
    
    style = [('BACKGROUND', (0, 0), (-1, 0), '#4472C4'),
             ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
             ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
             ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
             ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
             ('FONT', (0, 0), (-1, -1), 'Helvetica', 12)]

    t.setStyle(style)

    cell_height = 20 
    table_height = (len(data) + 1) * cell_height

    y = 47
    y2 = (table_height/20 - 1) * 2
    table_y = 765 - 60 - table_height + (y-y2)
    
    t.wrapOn(p, 0, 0)
    t.drawOn(p, 60, table_y)

    p.save()
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name=f'Report.pdf')

@app.route('/make-pdf-voltage', methods=['POST'])
def make_pdf_voltage():
    cust = request.form.get('cust')
    date = request.form.get('date')
    date_obj = datetime.strptime(date, "%Y-%m-%d")

    formatted_date = date_obj.strftime("%d %B %Y")
    data = CustUnbalance.query.filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).all()
    total_van_vbn_count = db.session.query(func.sum(CustUnbalance.van_vbn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_van_vcn_count = db.session.query(func.sum(CustUnbalance.van_vcn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vbn_vcn_count = db.session.query(func.sum(CustUnbalance.vbn_vcn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vab_vbc_count = db.session.query(func.sum(CustUnbalance.vab_vbc)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vab_vca_count = db.session.query(func.sum(CustUnbalance.vab_vca)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vbc_vca_count = db.session.query(func.sum(CustUnbalance.vbc_vca)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_unbalance_count = db.session.query(func.sum(CustUnbalance.unbalance_count)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    
    chunk_size = 32
    chunks = [data[i:i + chunk_size] for i in range(0, len(data), chunk_size)]
    
    p.setFont("Helvetica-Bold", 16)
    p.drawString(60, 785, f"{cust} Daily Report {formatted_date}")  
    p.setFont("Helvetica", 10)
    p.drawString(60, 765, f"Total Van & Vbn Unbalanced: {total_van_vbn_count}")
    p.drawString(60, 750, f"Total Van & Vcn Unbalanced: {total_van_vcn_count}")
    p.drawString(60, 735, f"Total Vbn & Vcn Unbalanced: {total_vbn_vcn_count}")
    p.drawString(60, 720, f"Total Vab & Vbc Unbalanced: {total_vab_vbc_count}")
    p.drawString(60, 705, f"Total Vab & Vca Unbalanced: {total_vab_vca_count}")
    p.drawString(60, 690, f"Total Vbc & Vca Unbalanced: {total_vbc_vca_count}")
    p.drawString(60, 675, f"Total Unbalanced: {total_unbalance_count}")
    
    for chunk in chunks:
        
        table_data = [['Van', 'Vbn', 'Vcn', 'Vab', 'Vbc', 'Vca', 'Van & Vbn', 'Van & Vcn', 'Vbn & Vcn', 'Vab & Vbc', 'Vab & Vca', 'Vbc & Vca']]
        # table_data = [['Van', 'Vbn', 'Vcn', 'Vab', 'Vbc', 'Vca', 'Unbalance Count']]
        for row in chunk:
            # table_data.append([row.van, row.vbn, row.vcn, row.vab, row.vbc, row.vca, row.unbalance_count])
            table_data.append([row.van, row.vbn, row.vcn, row.vab, row.vbc, row.vca, row.van_vbn, row.van_vcn, row.vbn_vcn, row.vab_vbc, row.vab_vca, row.vbc_vca])

        col_widths = [39, 39, 39, 39, 39, 39, 40, 40, 40, 40, 40, 40]
        # col_widths = [60, 60, 60, 60, 60, 60, 115]
        t = Table(table_data, colWidths=col_widths)
        
        style = [('BACKGROUND', (0, 0), (-1, 0), '#4472C4'),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                ('FONT', (0, 0), (-1, -1), 'Helvetica', 8)]

        t.setStyle(style)

        cell_height = 20 
        table_height = (len(chunk) + 1) * cell_height
        
        y = (4.54545*len(chunk))+3.5455
        
        table_y = 675 - 30 - table_height + (y)
        
        
        t.wrapOn(p, 0, 0)
        t.drawOn(p, 60, table_y)
        p.showPage()

    p.save()
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name=f'Report.pdf')

if __name__ == '__main__':
    app.run(debug=True)
