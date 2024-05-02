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
from reportlab.lib.pagesizes import landscape

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
    van_vbn = db.Column(db.Float)
    van_vcn = db.Column(db.Float)
    vbn_vcn = db.Column(db.Float)
    vab_vbc = db.Column(db.Float)
    vab_vca = db.Column(db.Float)
    vbc_vca = db.Column(db.Float)
    vbc_vca_unbalance = db.Column(db.Integer)
    van_vbn_unbalance = db.Column(db.Integer)
    van_vcn_unbalance = db.Column(db.Integer)
    vbn_vcn_unbalance = db.Column(db.Integer)
    vab_vbc_unbalance = db.Column(db.Integer)
    vab_vca_unbalance = db.Column(db.Integer)
    vbc_vca_unbalance = db.Column(db.Integer)
    van_status = db.Column(db.String(20))
    vbn_status = db.Column(db.String(20))
    vcn_status = db.Column(db.String(20))
    # van_vbn = db.Column(db.String(10))
    # van_vcn = db.Column(db.String(10))
    # vbn_vcn = db.Column(db.String(10))
    # vab_vbc = db.Column(db.String(10))
    # vab_vca = db.Column(db.String(10))
    # vbc_vca = db.Column(db.String(10))
    unbalance_count = db.Column(db.Integer)
    
    freq = db.Column(db.Float)
    oiltemp = db.Column(db.Float)
    wtitemp1 = db.Column(db.Float)
    wtitemp2 = db.Column(db.Float)
    wtitemp3 = db.Column(db.Float)
    pfa = db.Column(db.Float)
    pfb = db.Column(db.Float)
    pfc = db.Column(db.Float)
    ia = db.Column(db.Float)
    ib = db.Column(db.Float)
    ic = db.Column(db.Float)
    ineutral = db.Column(db.Float)
    bustemp1 = db.Column(db.Float)
    bustemp2 = db.Column(db.Float)
    bustemp3 = db.Column(db.Float)
    press = db.Column(db.Float)
    
    freq_status = db.Column(db.String(20))
    oiltemp_status = db.Column(db.String(20))
    wtitemp1_status = db.Column(db.String(20))
    wtitemp2_status = db.Column(db.String(20))
    wtitemp3_status = db.Column(db.String(20))
    pfa_status = db.Column(db.String(20))
    pfb_status = db.Column(db.String(20))
    pfc_status = db.Column(db.String(20))
    ia_status = db.Column(db.String(20))
    ib_status = db.Column(db.String(20))
    ic_status = db.Column(db.String(20))
    ineutral_status = db.Column(db.String(20))
    bustemp1_status = db.Column(db.String(20))
    bustemp2_status = db.Column(db.String(20))
    bustemp3_status = db.Column(db.String(20))
    press_status = db.Column(db.String(20))
    
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
    print(date)
    print(cust)

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
    cust = request.args.get('cust')
    data = CustUnbalance.query.filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).all()
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
    offline = request.form.get('offline')
    uploaded_files = request.files.getlist('excel_files')
    data = []
    

    existing_data = CustTemp.query.filter_by(nama=nama, date=date).first()
    if existing_data:
        return jsonify({"message": "existed"})
    
    if offline == '1':
        new = CustTemp(date=date, nama=nama, average_temp=0, safe_percentage=0)
        db.session.add(new)
        db.session.commit()
        return jsonify({"message": "success", "date": date})
        
    else:
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
            # df = pd.read_excel(excel_file, usecols="H,I,Q,R,Z,AA")
            df = pd.read_excel(excel_file, usecols="B,C,D,E,F,H,I,J,K,Q,R,S,T,Z,AA,AB,AC,AN,AO,AZ,BA,BB")

            for _, row in df.iterrows():
                
                # Unbalance
                if round(abs(row['Van'] - row['Vbn']))  > 15:
                    van_vbn = round(abs(row['Van'] - row['Vbn'])) 
                    van_vbn_unbalance = 1
                else:
                    van_vbn = round(abs(row['Van'] - row['Vbn'])) 
                    van_vbn_unbalance = 0
                if round(abs(row['Van'] - row['Vcn']))  > 15:
                    van_vcn = round(abs(row['Van'] - row['Vcn'])) 
                    van_vcn_unbalance = 1
                else:
                    van_vcn = round(abs(row['Van'] - row['Vcn'])) 
                    van_vcn_unbalance = 0
                if round(abs(row['Vbn'] - row['Vcn']))  > 15:
                    vbn_vcn = round(abs(row['Vbn'] - row['Vcn'])) 
                    vbn_vcn_unbalance = 1
                else:
                    vbn_vcn = round(abs(row['Vbn'] - row['Vcn'])) 
                    vbn_vcn_unbalance = 0
                if round(abs(row['Vab'] - row['Vbc']))  > 15:
                    vab_vbc = round(abs(row['Vab'] - row['Vbc'])) 
                    vab_vbc_unbalance = 1
                else:
                    vab_vbc = round(abs(row['Vab'] - row['Vbc'])) 
                    vab_vbc_unbalance = 0
                if round(abs(row['Vab'] - row['Vca']))  > 15:
                    vab_vca = round(abs(row['Vab'] - row['Vca'])) 
                    vab_vca_unbalance = 1
                else:
                    vab_vca = round(abs(row['Vab'] - row['Vca'])) 
                    vab_vca_unbalance = 0
                if round(abs(row['Vbc'] - row['Vca']))  > 15:
                    vbc_vca = round(abs(row['Vbc'] - row['Vca'])) 
                    vbc_vca_unbalance = 1
                else:
                    vbc_vca = round(abs(row['Vbc'] - row['Vca'])) 
                    vbc_vca_unbalance = 0
                    
                # Under / Over Voltage
                if (row['Van'] <= 207):
                    van_status = 'Under Voltage'
                elif (row['Van'] >= 241.5):
                    van_status = 'Over Voltage'
                else:
                    van_status = 'Normal'
                if (row['Vbn'] <= 207):
                    vbn_status = 'Under Voltage'
                elif (row['Vbn'] >= 241.5):
                    van_status = 'Over Voltage'
                else:
                    vbn_status = 'Normal'
                if (row['Vcn'] <= 207):
                    vcn_status = 'Under Voltage'
                elif (row['Vcn'] >= 241.5):
                    van_status = 'Over Voltage'
                else:
                    vcn_status = 'Normal'
                    
                # Freq
                if (row['Freq'] < 50):
                    freq_status = 'Low Alarm'
                elif (row['Freq'] > 50):
                    freq_status = 'High Alarm'
                else:
                    freq_status = 'Normal'
                    
                # Top Oil
                
                if (row['OilTemp'] > 70):
                    oiltemp_status = 'High Alarm'
                else:
                    oiltemp_status = 'Normal'
                
                # WTI

                
                if (row['WTITemp1'] > 88):
                    wtitemp1_status = 'High Alarm'
                else:
                    wtitemp1_status = 'Normal'
                    
                if (row['WTITemp2'] > 88):
                    wtitemp2_status = 'High Alarm'
                else:
                    wtitemp2_status = 'Normal'
                    
                if (row['WTITemp3'] > 88):
                    wtitemp3_status = 'High Alarm'
                else:
                    wtitemp3_status = 'Normal'
                
                # PF

                
                if (row['PFa'] < 0.75):
                    pfa_status = 'Low Alarm'
                else:
                    pfa_status = 'Normal'
                    
                if (row['PFb'] < 0.75):
                    pfb_status = 'Low Alarm'
                else:
                    pfb_status = 'Normal'
            
                if (row['PFc'] < 0.75):
                    pfc_status = 'Low Alarm'
                else:
                    pfc_status = 'Normal'
                
                # Current

                
                if (row['Ia'] > row['Ia']*1.1):
                    ia_status = 'High Alarm'
                else:
                    ia_status = 'Normal'
                
                if (row['Ib'] > row['Ib']*1.1):
                    ib_status = 'High Alarm'
                else:
                    ib_status = 'Normal'
                    
                if (row['Ic'] > row['Ic']*1.1):
                    ic_status = 'High Alarm'
                else:
                    ic_status = 'Normal'
                
                # Neutral Current

                
                if (row['Ineutral'] > 125):
                    ineutral_status = 'High Alarm'
                else:
                    ineutral_status = 'Normal'
                
                # Busbar Temp

                
                if (row['BusTemp1'] > 54):
                    bustemp1_status = 'High Alarm'
                else:
                    bustemp1_status = 'Normal'
                
                if (row['BusTemp2'] > 54):
                    bustemp2_status = 'High Alarm'
                else:
                    bustemp2_status = 'Normal'
                    
                if (row['BusTemp3'] > 54):
                    bustemp3_status = 'High Alarm'
                else:
                    bustemp3_status = 'Normal'
                
                # Tank Pressure
                if (row['Press'] > 0.5):
                    press_status = 'High Alarm'
                else:
                    press_status = 'Normal'
                    
                # diff = abs(row['Van'] - row['Vbn']) + abs(row['Van'] - row['Vcn']) + abs(row['Vbn'] - row['Vcn'])
                unbalance_count = van_vbn_unbalance + van_vcn_unbalance + vbn_vcn_unbalance + vab_vbc_unbalance + vab_vca_unbalance + vbc_vca_unbalance
                
                new = CustUnbalance(date=date, nama=nama, van=row['Van'], vbn=row['Vbn'], vcn=row['Vcn'], vab=row['Vab'], vbc=row['Vbc'], vca=row['Vca'],
                                    van_vbn=van_vbn, van_vcn=van_vcn, vbn_vcn=vbn_vcn, vab_vbc=vab_vbc, vab_vca=vab_vca, vbc_vca=vbc_vca,
                                    van_vbn_unbalance=van_vbn_unbalance, van_vcn_unbalance=van_vcn_unbalance, vbn_vcn_unbalance=vbn_vcn_unbalance,
                                    vab_vbc_unbalance=vab_vbc_unbalance, vab_vca_unbalance=vab_vca_unbalance, vbc_vca_unbalance=vbc_vca_unbalance,
                                    van_status=van_status, vbn_status=vbn_status, vcn_status=vcn_status,
                                    freq_status=freq_status, oiltemp_status=oiltemp_status,
                                    wtitemp1_status=wtitemp1_status, wtitemp2_status=wtitemp2_status, wtitemp3_status=wtitemp3_status,
                                    pfa_status=pfa_status, pfb_status=pfb_status, pfc_status=pfc_status,
                                    ia_status=ia_status, ib_status=ib_status, ic_status=ic_status, ineutral_status=ineutral_status,
                                    bustemp1_status=bustemp1_status, bustemp2_status=bustemp2_status, bustemp3_status=bustemp3_status, press_status=press_status,
                                    freq = row['Freq'], oiltemp = row['OilTemp'], wtitemp1 = row['WTITemp1'], wtitemp2 = row['WTITemp2'], wtitemp3 = row['WTITemp3'],
                                    pfa = row['PFa'], pfb = row['PFb'], pfc = row['PFc'],
                                    ia = row['Ia'], ib = row['Ib'], ic = row['Ic'],
                                    ineutral = row['Ineutral'], bustemp1 = row['BusTemp1'], bustemp2 = row['BusTemp2'], bustemp3 = row['BusTemp3'],
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
    
    total_van_vbn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.van_vbn_unbalance == 1)).scalar()
    total_van_vcn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.van_vcn_unbalance == 1)).scalar()
    total_vbn_vcn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbn_vcn_unbalance == 1)).scalar()
    total_vab_vbc_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vab_vbc_unbalance == 1)).scalar()
    total_vab_vca_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vab_vca_unbalance == 1)).scalar()
    total_vbc_vca_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbc_vca_unbalance == 1)).scalar()
    
    total_van_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbc_vca_unbalance == 'Under Voltage')).scalar()
    total_vbn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbc_vca_unbalance == 'Under Voltage')).scalar()
    total_vcn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbc_vca_unbalance == 'Under Voltage')).scalar()
    total_v_count = total_van_count + total_vbn_count + total_vcn_count
    

    # total_van_vbn_count = db.session.query(func.sum(CustUnbalance.van_vbn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_van_vcn_count = db.session.query(func.sum(CustUnbalance.van_vcn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vbn_vcn_count = db.session.query(func.sum(CustUnbalance.vbn_vcn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vab_vbc_count = db.session.query(func.sum(CustUnbalance.vab_vbc)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vab_vca_count = db.session.query(func.sum(CustUnbalance.vab_vca)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vbc_vca_count = db.session.query(func.sum(CustUnbalance.vbc_vca)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_unbalance_count = db.session.query(func.sum(CustUnbalance.unbalance_count)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=landscape(A4))
    
    chunk_size = 30
    chunks = [data[i:i + chunk_size] for i in range(0, len(data), chunk_size)]
    
    p.line(80, 575, 760, 575)
    
    p.setFont("Helvetica-Bold", 25)
    text_width = p.stringWidth("VOLTAGE DAILY REPORT", "Helvetica-Bold", 25)
    x_center = (landscape(A4)[0] - text_width) / 2
    # print(x_center)
    p.drawString(x_center, 540, "VOLTAGE DAILY REPORT")
    
    p.line(80, 525, 760, 525)
    
    # Name
    p.setFillColor(colors.lightgrey)
    p.rect(80, 480, 150, 32, fill=1)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 16)
    p.drawString(85, 490, "NAME")
    
    p.setFillColor(colors.lightgrey)
    p.rect(230, 480, 530, 32, fill=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 16)
    p.drawString(235, 490, f"{cust}")
    
    # Date
    p.setFillColor(colors.lightgrey)
    p.rect(80, 440, 150, 32, fill=1)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 16)
    p.drawString(85, 450, "DATE")
    
    p.setFillColor(colors.lightgrey)
    p.rect(230, 440, 530, 32, fill=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 16)
    p.drawString(235, 450, f"{formatted_date}")
    
    # Summary
    p.setFillColor(colors.lightgrey)
    p.rect(80, 400, 680, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 15)
    text_width = p.stringWidth("SUMMARY", "Helvetica", 15)
    x_center = (landscape(A4)[0] - text_width) / 2
    p.drawString(x_center, 410, "SUMMARY")
    
    start = 365
    
    # Van & Vbn
    p.setFillColor(colors.lightgrey)
    p.rect(80, start, 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start+10, "Total Van & Vbn Unbalance")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start, 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start+10, f"{total_van_vbn_count}")

    # Van & Vcn
    p.setFillColor(colors.lightgrey)
    p.rect(80, start-35, 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start-35+10, "Total Van & Vcn Unbalance")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start-35, 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start-35+10, f"{total_van_vcn_count}")

    # Total Vbn & Vcn Unbalance
    p.setFillColor(colors.lightgrey)
    p.rect(80, start-(35*2), 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start-(35*2)+10, "Total Vbn & Vcn Unbalance")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start-(35*2), 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start-(35*2)+10, f"{total_vbn_vcn_count}")

    # Total Vab & Vbc Unbalance
    p.setFillColor(colors.lightgrey)
    p.rect(80, start-(35*3), 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start-(35*3)+10, "Total Vab & Vbc Unbalance")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start-(35*3), 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start-(35*3)+10, f"{total_vab_vbc_count}")

    # Total Vab & Vca Unbalance
    p.setFillColor(colors.lightgrey)
    p.rect(80, start-(35*4), 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start-(35*4)+10, "Total Vab & Vca Unbalance")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start-(35*4), 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start-(35*4)+10, f"{total_vab_vca_count}")

    # Total Vbc & Vca Unbalance
    p.setFillColor(colors.lightgrey)
    p.rect(80, start-(35*5), 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start-(35*5)+10, "Total Vbc & Vca Unbalance")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start-(35*5), 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start-(35*5)+10, f"{total_vbc_vca_count}")

    # Total Unbalance
    p.setFillColor(colors.lightgrey)
    p.rect(80, start-(35*6), 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start-(35*6)+10, "Total Unbalance")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start-(35*6), 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start-(35*6)+10, f"{total_unbalance_count}")

    # Total Van Count
    p.setFillColor(colors.lightgrey)
    p.rect(80, start-(35*7), 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start-(35*7)+10, "Total Under Voltage Van")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start-(35*7), 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start-(35*7)+10, f"{total_van_count}")

    # Total Vbn Count
    p.setFillColor(colors.lightgrey)
    p.rect(80, start-(35*8), 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start-(35*8)+10, "Total Under Voltage Vbn")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start-(35*8), 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start-(35*8)+10, f"{total_vbn_count}")

    # Total Vcn Count
    p.setFillColor(colors.lightgrey)
    p.rect(80, start-(35*9), 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start-(35*9)+10, "Total Under Voltage Vcn")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start-(35*9), 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start-(35*9)+10, f"{total_vcn_count}")

    # Total V Count
    p.setFillColor(colors.lightgrey)
    p.rect(80, start-(35*10), 510, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(85, start-(35*10)+10, "Total Under Voltage")

    p.setFillColor(colors.lightgrey)
    p.rect(595, start-(35*10), 165, 30, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 12)
    p.drawString(675, start-(35*10)+10, f"{total_v_count}")

    p.showPage()

    
    # p.line(60, 800, 535, 800)

    # p.setFont("Helvetica-Bold", 25)
    # text_width = p.stringWidth("VOLTAGE DAILY REPORT", "Helvetica-Bold", 25)
    # x_center = (A4[0] - text_width) / 2
    # p.drawString(x_center, 765, "VOLTAGE DAILY REPORT")

    # p.line(60, 750, 535, 750)
    
    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 705, 100, 30, fill=1)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 14)
    # p.drawString(65, 715, "NAME")
    # p.setFont("Helvetica", 14)
    
    # p.setFillColor(colors.lightgrey)
    # p.rect(160, 705, 375, 30, fill=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 14)
    # p.drawString(165, 715, f"{cust}")
    
    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 665, 100, 30, fill=1)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 14)
    # p.drawString(65, 675, "DATE")
    # p.setFont("Helvetica", 14)
    
    # p.setFillColor(colors.lightgrey)
    # p.rect(160, 665, 375, 30, fill=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 14)
    # p.drawString(165, 675, f"{formatted_date}")

    # # p.setFillColor(colors.lightgrey)
    # # p.rect(60, 705, 225, 30, fill=1, stroke=0)
    # # p.setFillColor(colors.black)
    # # p.setFont("Helvetica", 9)
    # # p.drawString(65, 725, "NAME")
    # # p.setFont("Helvetica", 14)
    # # p.drawString(65, 710, f"{cust}")
    
    # # p.setFillColor(colors.lightgrey)
    # # p.rect(310, 705, 225, 30, fill=1, stroke=0)
    # # p.setFillColor(colors.black)
    # # p.setFont("Helvetica", 9)
    # # p.drawString(315, 725, "DATE")
    # # p.setFont("Helvetica", 14)
    # # p.drawString(315, 710, f"{formatted_date}")
    
    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 605, 475, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 15)
    # text_width = p.stringWidth("SUMMARY", "Helvetica", 15)
    # x_center = (A4[0] - text_width) / 2
    # p.drawString(x_center, 615, "SUMMARY")
    
    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 570, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 580, "Total Van & Vbn Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 570, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 580, f"{total_van_vbn_count}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 535, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 545, "Total Van & Vcn Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 535, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 545, f"{total_van_vcn_count}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 500, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 510, "Total Vbn & Vcn Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 500, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 510, f"{total_vbn_vcn_count}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 465, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 475, "Total Vab & Vbc Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 465, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 475, f"{total_vab_vbc_count}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 430, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 440, "Total Vab & Vca Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 430, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 440, f"{total_vab_vca_count}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 395, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 405, "Total Vbc & Vca Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 395, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 405, f"{total_vbc_vca_count}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 360, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 370, "Total Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 360, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 370, f"{total_unbalance_count}")
    
    
    for chunk in chunks:
        table_data = [['Van', 'Vbn', 'Vcn',
                       'Vab', 'Vbc', 'Vca',
                       'Van & Vbn', 'Van & Vcn', 'Vbn & Vcn',
                       'Vab & Vbc', 'Vab & Vca', 'Vbc & Vca',
                       'Van Status', 'Vbn Status', 'Vcn Status']]
        # table_data = [['Van', 'Vbn', 'Vcn',
        #                'Vab', 'Vbc', 'Vca',
        #                'Freq', 'Oil Temp',
        #                'WTI Temp 1', 'WTI Temp 2', 'WTI Temp 3',
        #                'PF a', 'PF b', 'PF c',
        #                'Current A', 'Current B', 'Current C',
        #                'Neutral Current',
        #                'Busbar Temp 1', 'Busbar Temp 2', 'Busbar Temp 3',
        #                'Pressure',
        #                'Van & Vbn', 'Van & Vcn', 'Vbn & Vcn',
        #                'Vab & Vbc', 'Vab & Vca', 'Vbc & Vca',
        #                'Van Status', 'Vbn Status', 'Vcn Status']]
        
        for row in chunk:
            table_data.append([row.van, row.vbn, row.vcn,
                               row.vab, row.vbc, row.vca, row.van_vbn, row.van_vcn,
                               row.vbn_vcn, row.vab_vbc, row.vab_vca, row.vbc_vca,
                               row.van_status, row.vbn_status, row.vcn_status])

        col_widths = [39, 39, 39, 39, 39, 39, 40, 40, 40, 40, 40, 40, 70, 70, 70]
        # col_widths = [39, 39, 39, 39, 39, 39, 40, 40, 40, 40, 40, 40, 40, 40, 40]
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
        
        table_y = 550 - table_height + (y)
        
        
        t.wrapOn(p, 0, 0)
        t.drawOn(p, 80, table_y)
        p.showPage()

    p.save()
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name=f'Report.pdf')

if __name__ == '__main__':
    app.run(debug=True)
