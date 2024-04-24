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
    # van_vbn = db.Column(db.Integer)
    # van_vcn = db.Column(db.Integer)
    # vbn_vcn = db.Column(db.Integer)
    # vab_vbc = db.Column(db.Integer)
    # vab_vca = db.Column(db.Integer)
    # vbc_vca = db.Column(db.Integer)
    van_vbn = db.Column(db.String(10))
    van_vcn = db.Column(db.String(10))
    vbn_vcn = db.Column(db.String(10))
    vab_vbc = db.Column(db.String(10))
    vab_vca = db.Column(db.String(10))
    vbc_vca = db.Column(db.String(10))
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
                # if abs(row['Van'] - row['Vbn']) > 15:
                #     van_vbn = 1
                # else:
                #     van_vbn = 0
                # if abs(row['Van'] - row['Vcn']) > 15:
                #     van_vcn = 1
                # else:
                #     van_vcn = 0
                # if abs(row['Vbn'] - row['Vcn']) > 15:
                #     vbn_vcn = 1
                # else:
                #     vbn_vcn = 0
                # if abs(row['Vab'] - row['Vbc']) > 15:
                #     vab_vbc = 1
                # else:
                #     vab_vbc = 0
                # if abs(row['Vab'] - row['Vca']) > 15:
                #     vab_vca = 1
                # else:
                #     vab_vca = 0
                # if abs(row['Vbc'] - row['Vca']) > 15:
                #     vbc_vca = 1
                # else:
                #     vbc_vca = 0
                
                if abs(row['Van'] - row['Vbn']) > 15:
                    van_vbn = 'Unbalance'
                    van_vbn2 = 1
                else:
                    van_vbn = 'Balance'
                    van_vbn2 = 0
                if abs(row['Van'] - row['Vcn']) > 15:
                    van_vcn = 'Unbalance'
                    van_vcn2 = 1
                else:
                    van_vcn = 'Balance'
                    van_vcn2 = 0
                if abs(row['Vbn'] - row['Vcn']) > 15:
                    vbn_vcn = 'Unbalance'
                    vbn_vcn2 = 1
                else:
                    vbn_vcn = 'Balance'
                    vbn_vcn2 = 0
                if abs(row['Vab'] - row['Vbc']) > 15:
                    vab_vbc = 'Unbalance'
                    vab_vbc2 = 1
                else:
                    vab_vbc = 'Balance'
                    vab_vbc2 = 0
                if abs(row['Vab'] - row['Vca']) > 15:
                    vab_vca = 'Unbalance'
                    vab_vca2 = 1
                else:
                    vab_vca = 'Balance'
                    vab_vca2 = 0
                if abs(row['Vbc'] - row['Vca']) > 15:
                    vbc_vca = 'Unbalance'
                    vbc_vca2 = 1
                else:
                    vbc_vca = 'Balance'
                    vbc_vca2 = 0
                    
                # diff = abs(row['Van'] - row['Vbn']) + abs(row['Van'] - row['Vcn']) + abs(row['Vbn'] - row['Vcn'])
                unbalance_count = van_vbn2 + van_vcn2 + vbn_vcn2 + vab_vbc2 + vab_vca2 + vbc_vca2
                
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
    
    total_van_vbn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.van_vbn == 'Unbalance')).scalar()
    total_van_vcn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.van_vcn == 'Unbalance')).scalar()
    total_vbn_vcn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbn_vcn == 'Unbalance')).scalar()
    total_vab_vbc_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vab_vbc == 'Unbalance')).scalar()
    total_vab_vca_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vab_vca == 'Unbalance')).scalar()
    total_vbc_vca_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbc_vca == 'Unbalance')).scalar()

    # total_van_vbn_count = db.session.query(func.sum(CustUnbalance.van_vbn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_van_vcn_count = db.session.query(func.sum(CustUnbalance.van_vcn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vbn_vcn_count = db.session.query(func.sum(CustUnbalance.vbn_vcn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vab_vbc_count = db.session.query(func.sum(CustUnbalance.vab_vbc)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vab_vca_count = db.session.query(func.sum(CustUnbalance.vab_vca)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vbc_vca_count = db.session.query(func.sum(CustUnbalance.vbc_vca)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_unbalance_count = db.session.query(func.sum(CustUnbalance.unbalance_count)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    
    chunk_size = 45
    chunks = [data[i:i + chunk_size] for i in range(0, len(data), chunk_size)]
    
    p.line(60, 800, 535, 800)

    # Tulis "VOLTAGE DAILY REPORT" di tengah halaman
    p.setFont("Helvetica-Bold", 25)
    text_width = p.stringWidth("VOLTAGE DAILY REPORT", "Helvetica-Bold", 25)
    x_center = (A4[0] - text_width) / 2
    p.drawString(x_center, 765, "VOLTAGE DAILY REPORT")

    # Gambar garis horizontal lagi
    p.line(60, 750, 535, 750)
    
    p.setFillColor(colors.lightgrey)
    p.rect(60, 705, 100, 30, fill=1)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 14)
    p.drawString(65, 715, "NAME")
    p.setFont("Helvetica", 14)
    
    p.setFillColor(colors.lightgrey)
    p.rect(160, 705, 375, 30, fill=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 14)
    p.drawString(165, 715, f"{cust}")
    
    p.setFillColor(colors.lightgrey)
    p.rect(60, 665, 100, 30, fill=1)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 14)
    p.drawString(65, 675, "DATE")
    p.setFont("Helvetica", 14)
    
    p.setFillColor(colors.lightgrey)
    p.rect(160, 665, 375, 30, fill=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 14)
    p.drawString(165, 675, f"{formatted_date}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 705, 225, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 9)
    # p.drawString(65, 725, "NAME")
    # p.setFont("Helvetica", 14)
    # p.drawString(65, 710, f"{cust}")
    
    # p.setFillColor(colors.lightgrey)
    # p.rect(310, 705, 225, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 9)
    # p.drawString(315, 725, "DATE")
    # p.setFont("Helvetica", 14)
    # p.drawString(315, 710, f"{formatted_date}")
    
    p.setFillColor(colors.lightgrey)
    p.rect(60, 605, 475, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 15)
    text_width = p.stringWidth("SUMMARY", "Helvetica", 15)
    x_center = (A4[0] - text_width) / 2
    p.drawString(x_center, 615, "SUMMARY")
    
    # Kotak pertama di bawah kotak "SUMMARY"
    p.setFillColor(colors.lightgrey)
    p.rect(60, 570, 400, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(65, 580, "Total Van & Vbn Unbalance")

    # Kotak pertama di bawah kotak "SUMMARY"
    p.setFillColor(colors.lightgrey)
    p.rect(465, 570, 70, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(495, 580, f"{total_van_vbn_count}")

    # Kotak span berwarna abu-abu dengan tulisan "SUMMARY" di tengah kotak
    p.setFillColor(colors.lightgrey)
    p.rect(60, 535, 400, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(65, 545, "Total Van & Vcn Unbalance")

    # Kotak pertama di bawah kotak "SUMMARY"
    p.setFillColor(colors.lightgrey)
    p.rect(465, 535, 70, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(495, 545, f"{total_van_vcn_count}")

    # Kotak span berwarna abu-abu dengan tulisan "SUMMARY" di tengah kotak
    p.setFillColor(colors.lightgrey)
    p.rect(60, 500, 400, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(65, 510, "Total Vbn & Vcn Unbalance")

    # Kotak pertama di bawah kotak "SUMMARY"
    p.setFillColor(colors.lightgrey)
    p.rect(465, 500, 70, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(495, 510, f"{total_vbn_vcn_count}")

    # Kotak span berwarna abu-abu dengan tulisan "SUMMARY" di tengah kotak
    p.setFillColor(colors.lightgrey)
    p.rect(60, 465, 400, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(65, 475, "Total Vab & Vbc Unbalance")

    # Kotak pertama di bawah kotak "SUMMARY"
    p.setFillColor(colors.lightgrey)
    p.rect(465, 465, 70, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(495, 475, f"{total_vab_vbc_count}")

    # Kotak span berwarna abu-abu dengan tulisan "SUMMARY" di tengah kotak
    p.setFillColor(colors.lightgrey)
    p.rect(60, 430, 400, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(65, 440, "Total Vab & Vca Unbalance")

    # Kotak pertama di bawah kotak "SUMMARY"
    p.setFillColor(colors.lightgrey)
    p.rect(465, 430, 70, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(495, 440, f"{total_vab_vca_count}")

    # Kotak span berwarna abu-abu dengan tulisan "SUMMARY" di tengah kotak
    p.setFillColor(colors.lightgrey)
    p.rect(60, 395, 400, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(65, 405, "Total Vbc & Vca Unbalance")

    # Kotak pertama di bawah kotak "SUMMARY"
    p.setFillColor(colors.lightgrey)
    p.rect(465, 395, 70, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(495, 405, f"{total_vbc_vca_count}")

    # Kotak span berwarna abu-abu dengan tulisan "SUMMARY" di tengah kotak
    p.setFillColor(colors.lightgrey)
    p.rect(60, 360, 400, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(65, 370, "Total Unbalance")

    # Kotak pertama di bawah kotak "SUMMARY"
    p.setFillColor(colors.lightgrey)
    p.rect(465, 360, 70, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(495, 370, f"{total_unbalance_count}")
    
    
    
    # p.showPage()
    
    # p.line(60, 800, 575, 800)
    # p.setFont("Helvetica-Bold", 18)
    # p.drawString(60, 795, f"{cust}")  
    # p.setFont("Helvetica-Bold", 24)
    # p.drawString(60, 785, " VOLTAGE DAILY REPORT")  
    # p.setFont("Helvetica-Bold", 20)
    # p.drawString(60, 775, f"{formatted_date}")  
    # p.setFont("Helvetica", 12)
    # p.drawString(60, 765, f"Total Van & Vbn Unbalance: {total_van_vbn_count}")
    # p.drawString(60, 750, f"Total Van & Vcn Unbalance: {total_van_vcn_count}")
    # p.drawString(60, 735, f"Total Vbn & Vcn Unbalance: {total_vbn_vcn_count}")
    # p.drawString(60, 720, f"Total Vab & Vbc Unbalance: {total_vab_vbc_count}")
    # p.drawString(60, 705, f"Total Vab & Vca Unbalance: {total_vab_vca_count}")
    # p.drawString(60, 690, f"Total Vbc & Vca Unbalance: {total_vbc_vca_count}")
    # p.drawString(60, 675, f"Total Unbalance: {total_unbalance_count}")
    
    p.showPage()
    
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
                ('FONT', (0, 0), (-1, -1), 'Helvetica', 7.5)]

        t.setStyle(style)

        cell_height = 20 
        table_height = (len(chunk) + 1) * cell_height
        
        y = (4.54545*len(chunk))+3.5455
        
        table_y = 780 - table_height + (y)
        
        
        t.wrapOn(p, 0, 0)
        t.drawOn(p, 60, table_y)
        p.showPage()

    p.save()
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name=f'Report.pdf')

if __name__ == '__main__':
    app.run(debug=True)
