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
    date = db.Column(db.String(10), default='-')
    nama = db.Column(db.String(255))
    average_temp = db.Column(db.Float)
    safe_percentage = db.Column(db.Float)
    
    def __repr__(self):
        return f'<{self.nama}>'
    
class CustUnbalance(db.Model):
    id = db.Column(db.Integer, autoincrement=True, primary_key=True)
    date = db.Column(db.String(10), default='-')
    nama = db.Column(db.String(255))
    timestamp = db.Column(db.String(255))
    max_current = db.Column(db.Float)
    van = db.Column(db.Float)
    vbn = db.Column(db.Float)
    vcn = db.Column(db.Float)
    vab = db.Column(db.Float)
    vbc = db.Column(db.Float)
    vca = db.Column(db.Float)
    # van_vbn = db.Column(db.Float)
    # van_vcn = db.Column(db.Float)
    # vbn_vcn = db.Column(db.Float)
    # vab_vbc = db.Column(db.Float)
    # vab_vca = db.Column(db.Float)
    # vbc_vca = db.Column(db.Float)
    # vbc_vca_unbalance = db.Column(db.Integer)
    # van_vbn_unbalance = db.Column(db.Integer)
    # van_vcn_unbalance = db.Column(db.Integer)
    # vbn_vcn_unbalance = db.Column(db.Integer)
    # vab_vbc_unbalance = db.Column(db.Integer)
    # vab_vca_unbalance = db.Column(db.Integer)
    # vbc_vca_unbalance = db.Column(db.Integer)
    # van_status = db.Column(db.String(20))
    # vbn_status = db.Column(db.String(20))
    # vcn_status = db.Column(db.String(20))
    # van_vbn = db.Column(db.String(10), default='-')
    # van_vcn = db.Column(db.String(10), default='-')
    # vbn_vcn = db.Column(db.String(10), default='-')
    # vab_vbc = db.Column(db.String(10), default='-')
    # vab_vca = db.Column(db.String(10), default='-')
    # vbc_vca = db.Column(db.String(10), default='-')
    # unbalance_count = db.Column(db.Integer)
    
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
    
    under_voltage = db.Column(db.String(10), default='-')
    over_voltage = db.Column(db.String(10), default='-')
    under_freq = db.Column(db.String(10), default='-')
    over_freq = db.Column(db.String(10), default='-')
    high_thd_i = db.Column(db.String(10), default='-')
    high_thd_v = db.Column(db.String(10), default='-')
    oil_hightemp = db.Column(db.String(10), default='-')
    wti_hightemp = db.Column(db.String(10), default='-')
    low_pf = db.Column(db.String(10), default='-')
    over_current = db.Column(db.String(10), default='-')
    high_ineutral = db.Column(db.String(10), default='-')
    bus_hightemp = db.Column(db.String(10), default='-')
    high_press = db.Column(db.String(10), default='-')
    unbalance = db.Column(db.String(10), default='-')
    remarks = db.Column(db.String(255))
    
    van_under_voltage = db.Column(db.Integer)
    vbn_under_voltage = db.Column(db.Integer)
    vcn_under_voltage = db.Column(db.Integer)
    van_over_voltage = db.Column(db.Integer)
    vbn_over_voltage = db.Column(db.Integer)
    vcn_over_voltage = db.Column(db.Integer)
    high_thd_i1 = db.Column(db.Integer)
    high_thd_i2 = db.Column(db.Integer)
    high_thd_i3 = db.Column(db.Integer)
    high_thd_v1 = db.Column(db.Integer)
    high_thd_v2 = db.Column(db.Integer)
    high_thd_v3 = db.Column(db.Integer)
    wti_hightemp1 = db.Column(db.Integer)
    wti_hightemp2 = db.Column(db.Integer)
    wti_hightemp3 = db.Column(db.Integer)
    low_pfa = db.Column(db.Integer)
    low_pfb = db.Column(db.Integer)
    low_pfc = db.Column(db.Integer)
    over_currenta = db.Column(db.Integer)
    over_currentb = db.Column(db.Integer)
    over_currentc = db.Column(db.Integer)
    bus_hightemp1 = db.Column(db.Integer)
    bus_hightemp2 = db.Column(db.Integer)
    bus_hightemp3 = db.Column(db.Integer)
    van_vbn_unbalance = db.Column(db.Integer)
    van_vcn_unbalance = db.Column(db.Integer)
    vbn_vcn_unbalance = db.Column(db.Integer)
    vab_vbc_unbalance = db.Column(db.Integer)
    vab_vca_unbalance = db.Column(db.Integer)
    vbc_vca_unbalance = db.Column(db.Integer)
    
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
                'unbalance_count': item.van_vbn_unbalance+item.van_vcn_unbalance+item.vbn_vcn_unbalance+item.vab_vbc_unbalance+item.vab_vca_unbalance+item.vbc_vca_unbalance,
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
        data = [row[0] for row in CustUnbalance.query.with_entities(CustUnbalance.nama).distinct().all()]
        
        return jsonify({"message": "success", "data": data})
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
    max_current = float(request.form.get('arusMax'))
    uploaded_files = request.files.getlist('excel_files')

    existing_data = CustUnbalance.query.filter_by(nama=nama, date=date).first()
    if existing_data:
        return jsonify({"message": "existed"})
    
    for index, excel_file in enumerate(uploaded_files):
        if excel_file.filename != '':
            # df = pd.read_excel(excel_file, usecols="H,I,Q,R,Z,AA")
            df = pd.read_excel(excel_file)

            for _, row in df.iterrows():
                unbalance_count = 0
                remarks = ''
                under_voltage = '-'
                over_voltage = '-'
                under_freq = '-'
                over_freq = '-'
                high_thd_i = '-'
                high_thd_v = '-'
                oil_hightemp = '-'
                wti_hightemp = '-'
                low_pf = '-'
                over_current = '-'
                high_ineutral = '-'
                bus_hightemp = '-'
                high_press = '-'
                unbalance = '-'
                # print(remarks)
                
                # Under Voltage
                if (row['Van'] < 207):
                    under_voltage = '✔'
                    remarks += str(row['Van']) + ', '
                    van_under_voltage = 1
                else:
                    van_under_voltage = 0
                    if under_voltage != '✔':
                        under_voltage = '-'
                # print(remarks)
                    
                if (row['Vbn'] < 207):
                    under_voltage = '✔'
                    remarks += str(row['Vbn']) + ', '
                    vbn_under_voltage = 1
                    
                else:
                    vbn_under_voltage = 0
                    if under_voltage != '✔':
                        under_voltage = '-'
                # print(remarks)
                    
                if (row['Vcn'] < 207):
                    under_voltage = '✔'
                    remarks += str(row['Vcn']) + ', '
                    vcn_under_voltage = 1
                    
                else:
                    vcn_under_voltage = 0
                    if under_voltage != '✔':
                        under_voltage = '-'
                # print(remarks)
                    
                # Over Voltage
                
                if (row['Van'] > 241.5):
                    over_voltage = '✔'
                    remarks += str(row['Van']) + ', '
                    van_over_voltage = 1
                else:
                    van_over_voltage = 0
                    if over_voltage != '✔':
                        over_voltage = '-'
                # print(remarks)
                
                if (row['Vbn'] > 241.5):
                    over_voltage = '✔'
                    remarks += str(row['Vbn']) + ', '
                    vbn_over_voltage = 1
                    
                else:
                    vbn_over_voltage = 0
                    if over_voltage != '✔':
                        over_voltage = '-'
                # print(remarks)
                    
                if (row['Vcn'] > 241.5):
                    over_voltage = '✔'
                    remarks += str(row['Vcn']) + ', '
                    vcn_over_voltage = 1
                    
                else:
                    vcn_over_voltage = 0
                    if over_voltage != '✔':
                        over_voltage = '-'
                # print(remarks)
                    
                # Freq
                if (row['Freq'] < 50-(50*0.05)):
                    under_freq = '✔'
                    over_freq = '-'
                    remarks += str(row['Freq']) + ', '
                    
                elif (row['Freq'] > 50+(50*0.05)):
                    over_freq = '✔'
                    under_freq = '-'
                    remarks += str(row['Freq']) + ', '
                    
                else:
                    if over_freq != '✔':
                        over_freq = '-'
                    if under_freq != '✔':
                        under_freq = '-'
                # print(remarks)
                    
                # THDI
                if (row['THDI1'] > 5):
                    high_thd_i = '✔'
                    remarks += str(row['THDI1']) + ', '
                    high_thd_i1 = 1
                else:
                    high_thd_i1 = 0
                    if high_thd_i != '✔':
                        high_thd_i = '-'
                # print(remarks)
                    
                if (row['THDI2'] > 5):
                    high_thd_i = '✔'
                    remarks += str(row['THDI2']) + ', '
                    high_thd_i2 = 1
                    
                else:
                    high_thd_i2 = 0
                    if high_thd_i != '✔':
                        high_thd_i = '-'
                # print(remarks)
                    
                if (row['THDI3'] > 5):
                    high_thd_i = '✔'
                    remarks += str(row['THDI3'] )+ ', '
                    high_thd_i3 = 1
                    
                else:
                    high_thd_i3 = 0
                    if high_thd_i != '✔':
                        high_thd_i = '-'
                # print(remarks)
                
                # THDV
                if (row['THDV1'] > 5):
                    high_thd_v = '✔'
                    remarks += str(row['THDV1']) + ', '
                    high_thd_v1 = 1
                    
                else:
                    high_thd_v1 = 0
                    if high_thd_v != '✔':
                        high_thd_v = '-'
                # print(remarks)
                    
                if (row['THDV2'] > 5):
                    high_thd_v = '✔'
                    remarks += str(row['THDV2'])+ ', '
                    high_thd_v2 = 1
                    
                else:
                    high_thd_v2 = 0
                    if high_thd_v != '✔':
                        high_thd_v = '-'
                # print(remarks)
                    
                if (row['THDV3'] > 5):
                    high_thd_v = '✔'
                    remarks += str(row['THDV3']) + ', '
                    high_thd_v3 = 1
                    
                else:
                    high_thd_v3 = 0
                    if high_thd_v != '✔':
                        high_thd_v = '-'
                # print(remarks)
                    
                # Top Oil
                if (row['OilTemp'] > 70):
                    oil_hightemp = '✔'
                    remarks += str(row['OilTemp']) + ', '
                    
                else:
                    if oil_hightemp != '✔':
                        oil_hightemp = '-'
                # print(remarks)
                
                # WTI
                
                if (row['WTITemp1'] > 88):
                    wti_hightemp = '✔'
                    remarks += str(row['WTITemp1']) + ', '
                    wti_hightemp1 = 1
                else:
                    wti_hightemp1 = 0
                    if wti_hightemp != '✔':
                        wti_hightemp = '-'
                # print(remarks)
                    
                if (row['WTITemp2'] > 88):
                    wti_hightemp = '✔'
                    remarks += str(row['WTITemp2']) + ', '
                    wti_hightemp2 = 1
                else:
                    wti_hightemp2 = 0
                    if wti_hightemp != '✔':
                        wti_hightemp = '-'
                # print(remarks)
                    
                if (row['WTITemp3'] > 88):
                    wti_hightemp = '✔'
                    remarks += str(row['WTITemp3']) + ', '
                    wti_hightemp3 = 1
                else:
                    wti_hightemp3 = 0
                    if wti_hightemp != '✔':
                        wti_hightemp = '-'
                # print(remarks)
                
                # PF

                if (row['PFa'] < 0.75):
                    low_pf = '✔'
                    remarks += str(row['PFa']) + ', '
                    low_pfa = 1
                else:
                    low_pfa = 0
                    if low_pf != '✔':
                        low_pf = '-'
                # print(remarks)
                    
                if (row['PFb'] < 0.75):
                    low_pf = '✔'
                    remarks += str(row['PFb']) + ', '
                    low_pfb = 1
                    
                else:
                    low_pfb = 0
                    if low_pf != '✔':
                        low_pf = '-'
                # print(remarks)
                    
                if (row['PFc'] < 0.75):
                    low_pf = '✔'
                    remarks += str(row['PFc']) + ', '
                    low_pfc = 1
                else:
                    low_pfc = 0
                    if low_pf != '✔':
                        low_pf = '-'
                # print(remarks)
                
                # Current
                if (row['Ia'] > max_current):
                    over_current = '✔'
                    remarks += str(row['Ia']) + ', '
                    over_currenta = 1
                else:
                    over_currenta = 0
                    if over_current != '✔':
                        over_current = '-'
                # print(remarks)
                
                if (row['Ib'] > max_current):
                    over_current = '✔'
                    remarks += str(row['Ib']) + ', '
                    over_currentb = 1
                    
                else:
                    over_currentb = 0
                    if over_current != '✔':
                        over_current = '-'
                # print(remarks)
                    
                if (row['Ic'] > max_current):
                    over_current = '✔'
                    remarks += str(row['Ic']) + ', '
                    over_currentc = 1
                    
                else:
                    over_currentc = 0
                    if over_current != '✔':
                        over_current = '-'
                # print(remarks)
                
                # Neutral Current
                if (row['Ineutral'] > 125):
                    high_ineutral = '✔'
                    remarks += str(row['Ineutral']) + ', '
                    
                else:
                    if high_ineutral != '✔':
                        high_ineutral = '-'
                # print(remarks)
                
                # Busbar Temp
                if (row['BusTemp1'] > 54):
                    bus_hightemp = '✔'
                    remarks += str(row['BusTemp1']) + ', '
                    bus_hightemp1 = 1
                else:
                    bus_hightemp1 = 0
                    if bus_hightemp != '✔':
                        bus_hightemp = '-'
                # print(remarks)
                
                if (row['BusTemp2'] > 54):
                    bus_hightemp = '✔'
                    remarks += str(row['BusTemp2']) + ', '
                    bus_hightemp2 = 1
                    
                else:
                    bus_hightemp2 = 0
                    if bus_hightemp != '✔':
                        bus_hightemp = '-'
                # print(remarks)
                    
                if (row['BusTemp3'] > 54):
                    bus_hightemp = '✔'
                    remarks += str(row['BusTemp3']) + ', '
                    bus_hightemp3 = 1
                    
                else:
                    bus_hightemp3 = 0
                    if bus_hightemp != '✔':
                        bus_hightemp = '-'
                # print(remarks)
                
                # Tank Pressure
                if (row['Press'] > 0.5):
                    high_press = '✔'
                    remarks += str(row['Press']) + ', '
                else:
                    if high_press != '✔':
                        high_press = '-'
                # print(remarks)
                    
                # Unbalance
                if round(abs(row['Van'] - row['Vbn']), 2)  > 15:
                    # van_vbn = round(abs(row['Van'] - row['Vbn'])) 
                    unbalance = '✔'
                    van_vbn_unbalance = 1
                    remarks += str(round(abs(row['Van'] - row['Vbn']), 2)) + ', '
                    
                else:
                    van_vbn_unbalance = 0
                    if unbalance != '✔':
                        unbalance = '-'
                # print(remarks)
                    
                if round(abs(row['Van'] - row['Vcn']), 2)  > 15:
                    # van_vcn = round(abs(row['Van'] - row['Vcn'])) 
                    unbalance = '✔'
                    van_vcn_unbalance = 1
                    remarks += str(round(abs(row['Van'] - row['Vcn']), 2)) + ', '
                    
                else:
                    van_vcn_unbalance = 0
                    if unbalance != '✔':
                        unbalance = '-'
                # print(remarks)
                    
                if round(abs(row['Vbn'] - row['Vcn']), 2)  > 15:
                    # vbn_vcn = round(abs(row['Vbn'] - row['Vcn'])) 
                    unbalance = '✔'
                    vbn_vcn_unbalance = 1
                    remarks += str(round(abs(row['Vbn'] - row['Vcn']), 2)) + ', '
                    
                else:
                    vbn_vcn_unbalance = 0
                    if unbalance != '✔':
                        unbalance = '-'
                # print(remarks)
                    
                if round(abs(row['Vab'] - row['Vbc']), 2)  > 15:
                    # vab_vbc = round(abs(row['Vab'] - row['Vbc'])) 
                    unbalance = '✔'
                    vab_vbc_unbalance = 1
                    remarks += str(round(abs(row['Vab'] - row['Vbc']), 2)) + ', '
                    
                else:
                    vab_vbc_unbalance = 0
                    if unbalance != '✔':
                        unbalance = '-'
                # print(remarks)
                    
                if round(abs(row['Vab'] - row['Vca']), 2)  > 15:
                    # vab_vca = round(abs(row['Vab'] - row['Vca'])) 
                    unbalance = '✔'
                    vab_vca_unbalance = 1
                    remarks += str(round(abs(row['Vab'] - row['Vca']), 2)) + ', '
                    
                else:
                    vab_vca_unbalance = 0
                    
                    if unbalance != '✔':
                        unbalance = '-'
                # print(remarks)
                    
                if round(abs(row['Vbc'] - row['Vca']), 2)  > 15:
                    # vbc_vca = round(abs(row['Vbc'] - row['Vca'])) 
                    unbalance = '✔'
                    vbc_vca_unbalance = 1
                    remarks += str(round(abs(row['Vbc'] - row['Vca']), 2))
                    
                else:
                    vbc_vca_unbalance = 0
                    
                    if unbalance != '✔':
                        unbalance = '-'
                # print(remarks)
                    
                remarks = remarks.rstrip(', ')
                # print(remarks)
                    
            
                # diff = abs(row['Van'] - row['Vbn']) + abs(row['Van'] - row['Vcn']) + abs(row['Vbn'] - row['Vcn'])
                # unbalance_count = van_vbn_unbalance + van_vcn_unbalance + vbn_vcn_unbalance + vab_vbc_unbalance + vab_vca_unbalance + vbc_vca_unbalance
                
                # new = CustUnbalance(date=date, nama=nama, van=row['Van'], vbn=row['Vbn'], vcn=row['Vcn'], vab=row['Vab'], vbc=row['Vbc'], vca=row['Vca'],
                #                     van_vbn=van_vbn, van_vcn=van_vcn, vbn_vcn=vbn_vcn, vab_vbc=vab_vbc, vab_vca=vab_vca, vbc_vca=vbc_vca,
                #                     van_vbn_unbalance=van_vbn_unbalance, van_vcn_unbalance=van_vcn_unbalance, vbn_vcn_unbalance=vbn_vcn_unbalance,
                #                     vab_vbc_unbalance=vab_vbc_unbalance, vab_vca_unbalance=vab_vca_unbalance, vbc_vca_unbalance=vbc_vca_unbalance,
                #                     van_status=van_status, vbn_status=vbn_status, vcn_status=vcn_status,
                #                     freq_status=freq_status, oiltemp_status=oiltemp_status,
                #                     wtitemp1_status=wtitemp1_status, wtitemp2_status=wtitemp2_status, wtitemp3_status=wtitemp3_status,
                #                     pfa_status=pfa_status, pfb_status=pfb_status, pfc_status=pfc_status,
                #                     ia_status=ia_status, ib_status=ib_status, ic_status=ic_status, ineutral_status=ineutral_status,
                #                     bustemp1_status=bustemp1_status, bustemp2_status=bustemp2_status, bustemp3_status=bustemp3_status, press_status=press_status,
                #                     freq = row['Freq'], oiltemp = row['OilTemp'], wtitemp1 = row['WTITemp1'], wtitemp2 = row['WTITemp2'], wtitemp3 = row['WTITemp3'],
                #                     pfa = row['PFa'], pfb = row['PFb'], pfc = row['PFc'],
                #                     ia = row['Ia'], ib = row['Ib'], ic = row['Ic'],
                #                     ineutral = row['Ineutral'], bustemp1 = row['BusTemp1'], bustemp2 = row['BusTemp2'], bustemp3 = row['BusTemp3'],
                #                     unbalance_count=unbalance_count)
                new = CustUnbalance(date=date, nama=nama, timestamp=row['timestamp'], max_current=max_current, van=row['Van'], vbn=row['Vbn'], vcn=row['Vcn'], vab=row['Vab'], vbc=row['Vbc'], vca=row['Vca'],
                                    freq = row['Freq'], oiltemp = row['OilTemp'], wtitemp1 = row['WTITemp1'], wtitemp2 = row['WTITemp2'], wtitemp3 = row['WTITemp3'],
                                    pfa = row['PFa'], pfb = row['PFb'], pfc = row['PFc'],
                                    ia = row['Ia'], ib = row['Ib'], ic = row['Ic'], press=row['Press'],
                                    ineutral = row['Ineutral'], bustemp1 = row['BusTemp1'], bustemp2 = row['BusTemp2'], bustemp3 = row['BusTemp3'],
                                    under_voltage=under_voltage, over_voltage=over_voltage, under_freq=under_freq, over_freq=over_freq,
                                    high_thd_i=high_thd_i, high_thd_v=high_thd_v, oil_hightemp=oil_hightemp, wti_hightemp=wti_hightemp,
                                    low_pf=low_pf, over_current=over_current, high_ineutral=high_ineutral,
                                    bus_hightemp=bus_hightemp, high_press=high_press,
                                    unbalance=unbalance, remarks=remarks,
                                    van_under_voltage=van_under_voltage, vbn_under_voltage=vbn_under_voltage, vcn_under_voltage=vcn_under_voltage, van_over_voltage=van_over_voltage,
                                    vbn_over_voltage=vbn_over_voltage, vcn_over_voltage=vcn_over_voltage,
                                    high_thd_i1=high_thd_i1, high_thd_i2=high_thd_i2, high_thd_i3=high_thd_i3,
                                    high_thd_v1=high_thd_v1, high_thd_v2=high_thd_v2, high_thd_v3=high_thd_v3,
                                    wti_hightemp1=wti_hightemp1, wti_hightemp2=wti_hightemp2, wti_hightemp3=wti_hightemp3,
                                    low_pfa=low_pfa, low_pfb=low_pfb, low_pfc=low_pfc,
                                    over_currenta=over_currenta, over_currentb=over_currentb, over_currentc=over_currentc,
                                    bus_hightemp1=bus_hightemp1, bus_hightemp2=bus_hightemp2, bus_hightemp3=bus_hightemp3,
                                    van_vbn_unbalance=van_vbn_unbalance, van_vcn_unbalance=van_vcn_unbalance,
                                    vbn_vcn_unbalance=vbn_vcn_unbalance, vab_vbc_unbalance=vab_vbc_unbalance,
                                    vab_vca_unbalance=vab_vca_unbalance, vbc_vca_unbalance=vbc_vca_unbalance
                                    )
                db.session.add(new)
                db.session.commit()
                
    data = [row[0] for row in CustUnbalance.query.with_entities(CustUnbalance.nama).distinct().all()]

    return jsonify({"message": "success", "date": date, "data": data})
    
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
    
    # total_van_vbn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.van_vbn_unbalance == 1)).scalar()
    # total_van_vcn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.van_vcn_unbalance == 1)).scalar()
    # total_vbn_vcn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbn_vcn_unbalance == 1)).scalar()
    # total_vab_vbc_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vab_vbc_unbalance == 1)).scalar()
    # total_vab_vca_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vab_vca_unbalance == 1)).scalar()
    # total_vbc_vca_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbc_vca_unbalance == 1)).scalar()
    
    # total_van_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbc_vca_unbalance == 'Under Voltage')).scalar()
    # total_vbn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbc_vca_unbalance == 'Under Voltage')).scalar()
    # total_vcn_count = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.vbc_vca_unbalance == 'Under Voltage')).scalar()
    # total_v_count = total_van_count + total_vbn_count + total_vcn_count
    
    # total_unbalance_count = db.session.query(func.sum(CustUnbalance.unbalance_count)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    
    #---
    

    # total_van_vbn_count = db.session.query(func.sum(CustUnbalance.van_vbn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_van_vcn_count = db.session.query(func.sum(CustUnbalance.van_vcn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vbn_vcn_count = db.session.query(func.sum(CustUnbalance.vbn_vcn)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vab_vbc_count = db.session.query(func.sum(CustUnbalance.vab_vbc)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vab_vca_count = db.session.query(func.sum(CustUnbalance.vab_vca)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_vbc_vca_count = db.session.query(func.sum(CustUnbalance.vbc_vca)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    
    # Total Under Voltage
    total_van_under_voltage = db.session.query(func.sum(CustUnbalance.van_under_voltage)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vbn_under_voltage = db.session.query(func.sum(CustUnbalance.vbn_under_voltage)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vcn_under_voltage = db.session.query(func.sum(CustUnbalance.vcn_under_voltage)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_under_voltage = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.under_voltage == '✔')).scalar()
    
    # Total Over Voltage
    total_van_over_voltage = db.session.query(func.sum(CustUnbalance.van_over_voltage)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vbn_over_voltage = db.session.query(func.sum(CustUnbalance.vbn_over_voltage)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vcn_over_voltage = db.session.query(func.sum(CustUnbalance.vcn_over_voltage)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_over_voltage = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.over_voltage == '✔')).scalar()
    
    # Total Under Freq
    total_under_freq = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.under_freq == '✔')).scalar()
    
    # Total Over Freq
    total_over_freq = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.over_freq == '✔')).scalar()
    
    # Total High THDI
    total_high_thd_i1 = db.session.query(func.sum(CustUnbalance.high_thd_i1)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_high_thd_i2 = db.session.query(func.sum(CustUnbalance.high_thd_i2)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_high_thd_i3 = db.session.query(func.sum(CustUnbalance.high_thd_i3)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_high_thd_i = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.high_thd_i == '✔')).scalar()
    total_high_thd_i = total_high_thd_i1 + total_high_thd_i2 + total_high_thd_i3
    
    # Total High THDV
    total_high_thd_v1 = db.session.query(func.sum(CustUnbalance.high_thd_v1)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_high_thd_v2 = db.session.query(func.sum(CustUnbalance.high_thd_v2)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_high_thd_v3 = db.session.query(func.sum(CustUnbalance.high_thd_v3)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_high_thd_v = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.high_thd_v == '✔')).scalar()
    total_high_thd_v = total_high_thd_v1 + total_high_thd_v2 + total_high_thd_v3
    
    # Total Oil High Temp
    total_oil_hightemp = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.oil_hightemp == '✔')).scalar()
    
    # Total WTI High Temp
    total_wti_hightemp1 = db.session.query(func.sum(CustUnbalance.wti_hightemp1)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_wti_hightemp2 = db.session.query(func.sum(CustUnbalance.wti_hightemp2)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_wti_hightemp3 = db.session.query(func.sum(CustUnbalance.wti_hightemp3)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_wti_hightemp = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.wti_hightemp == '✔')).scalar()
    total_wti_hightemp = total_wti_hightemp1 + total_wti_hightemp2 + total_wti_hightemp3
    
    # Total Low PF
    total_low_pfa = db.session.query(func.sum(CustUnbalance.low_pfa)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_low_pfb = db.session.query(func.sum(CustUnbalance.low_pfb)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_low_pfc = db.session.query(func.sum(CustUnbalance.low_pfc)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_low_pf = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.low_pf == '✔')).scalar()
    total_low_pf = total_low_pfa + total_low_pfb + total_low_pfb + total_low_pfc
    
    # Total Over Current
    total_over_currenta = db.session.query(func.sum(CustUnbalance.over_currenta)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_over_currentb = db.session.query(func.sum(CustUnbalance.over_currentb)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_over_currentc = db.session.query(func.sum(CustUnbalance.over_currentc)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_over_current = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.over_current == '✔')).scalar()
    total_over_current = total_over_currenta + total_over_currentb + total_over_currentc
    
    # Total High I Neutral
    total_high_ineutral = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.high_ineutral == '✔')).scalar()
    
    # Total Bus High Temp
    total_bus_hightemp1 = db.session.query(func.sum(CustUnbalance.bus_hightemp1)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_bus_hightemp2 = db.session.query(func.sum(CustUnbalance.bus_hightemp2)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_bus_hightemp3 = db.session.query(func.sum(CustUnbalance.bus_hightemp3)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    # total_bus_hightemp = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.bus_hightemp == '✔')).scalar()
    total_bus_hightemp = total_bus_hightemp1 + total_bus_hightemp2 + total_bus_hightemp3
    
    # Total High Press
    total_high_press = db.session.query(func.count(CustUnbalance.id)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust, CustUnbalance.high_press == '✔')).scalar()
    
    # Total Unbalance
    total_van_vbn_unbalance = db.session.query(func.sum(CustUnbalance.van_vbn_unbalance)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_van_vcn_unbalance = db.session.query(func.sum(CustUnbalance.van_vcn_unbalance)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vbn_vcn_unbalance = db.session.query(func.sum(CustUnbalance.vbn_vcn_unbalance)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vab_vbc_unbalance = db.session.query(func.sum(CustUnbalance.vab_vbc_unbalance)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vab_vca_unbalance = db.session.query(func.sum(CustUnbalance.vab_vca_unbalance)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_vbc_vca_unbalance = db.session.query(func.sum(CustUnbalance.vbc_vca_unbalance)).filter(and_(CustUnbalance.date == date, CustUnbalance.nama == cust)).scalar()
    total_unbalance = total_van_vbn_unbalance + total_van_vcn_unbalance + total_vbn_vcn_unbalance + total_vab_vbc_unbalance + total_vab_vca_unbalance + total_vbc_vca_unbalance
    
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    
    
    
    # p.line(80, 575, 760, 575)
    
    # p.setFont("Helvetica-Bold", 25)
    # text_width = p.stringWidth("VOLTAGE DAILY REPORT", "Helvetica-Bold", 25)
    # x_center = (landscape(A4)[0] - text_width) / 2
    # # print(x_center)
    # p.drawString(x_center, 540, "VOLTAGE DAILY REPORT")
    
    # p.line(80, 525, 760, 525)
    
    # # Name
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, 480, 150, 32, fill=1)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 16)
    # p.drawString(85, 490, "NAME")
    
    # p.setFillColor(colors.lightgrey)
    # p.rect(230, 480, 530, 32, fill=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 16)
    # p.drawString(235, 490, f"{cust}")
    
    # # Date
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, 440, 150, 32, fill=1)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 16)
    # p.drawString(85, 450, "DATE")
    
    # p.setFillColor(colors.lightgrey)
    # p.rect(230, 440, 530, 32, fill=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 16)
    # p.drawString(235, 450, f"{formatted_date}")
    
    # # Summary
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, 400, 680, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 15)
    # text_width = p.stringWidth("SUMMARY", "Helvetica", 15)
    # x_center = (landscape(A4)[0] - text_width) / 2
    # p.drawString(x_center, 410, "SUMMARY")
    
    # start = 365
    
    # # Van & Vbn
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start, 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start+10, "Total Van & Vbn Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start, 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start+10, f"{total_van_vbn_unbalance}")

    # # Van & Vcn
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start-35, 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start-35+10, "Total Van & Vcn Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start-35, 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start-35+10, f"{total_van_vcn_unbalance}")

    # # Total Vbn & Vcn Unbalance
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start-(35*2), 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start-(35*2)+10, "Total Vbn & Vcn Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start-(35*2), 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start-(35*2)+10, f"{total_vbn_vcn_unbalance}")

    # # Total Vab & Vbc Unbalance
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start-(35*3), 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start-(35*3)+10, "Total Vab & Vbc Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start-(35*3), 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start-(35*3)+10, f"{total_vab_vbc_unbalance}")

    # # Total Vab & Vca Unbalance
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start-(35*4), 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start-(35*4)+10, "Total Vab & Vca Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start-(35*4), 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start-(35*4)+10, f"{total_vab_vca_unbalance}")

    # # Total Vbc & Vca Unbalance
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start-(35*5), 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start-(35*5)+10, "Total Vbc & Vca Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start-(35*5), 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start-(35*5)+10, f"{total_vbc_vca_unbalance}")

    # # Total Unbalance
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start-(35*6), 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start-(35*6)+10, "Total Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start-(35*6), 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start-(35*6)+10, f"{total_unbalance}")

    # # Total Van Count
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start-(35*7), 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start-(35*7)+10, "Total Under Voltage Van")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start-(35*7), 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start-(35*7)+10, f"{total_van_under_voltage}")

    # # Total Vbn Count
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start-(35*8), 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start-(35*8)+10, "Total Under Voltage Vbn")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start-(35*8), 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start-(35*8)+10, f"{total_vbn_under_voltage}")

    # # Total Vcn Count
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start-(35*9), 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start-(35*9)+10, "Total Under Voltage Vcn")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start-(35*9), 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start-(35*9)+10, f"{total_vcn_under_voltage}")

    # # Total V Count
    # p.setFillColor(colors.lightgrey)
    # p.rect(80, start-(35*10), 510, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(85, start-(35*10)+10, "Total Under Voltage")

    # p.setFillColor(colors.lightgrey)
    # p.rect(595, start-(35*10), 165, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(675, start-(35*10)+10, f"{total_under_voltage}")


    

    # --------------------------------------------------------------------------------------------
    
    p.line(40, 820, 555, 820)

    p.setFont("Helvetica-Bold", 25)
    text_width = p.stringWidth("TMU EVENT REPORT", "Helvetica-Bold", 25)
    x_center = (A4[0] - text_width) / 2
    p.drawString(x_center, 793, "TMU EVENT REPORT")

    p.line(40, 785, 555, 785)
    
    p.setFillColor(colors.lightgrey)
    p.rect(40, 750, 100, 25, fill=1)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 14)
    p.drawString(45, 757.5, "NAME")
    p.setFont("Helvetica", 14)
    
    p.setFillColor(colors.lightgrey)
    p.rect(140, 750, 415, 25, fill=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 14)
    p.drawString(145, 757.5, f"{cust}")
    
    p.setFillColor(colors.lightgrey)
    p.rect(40, 720, 100, 25, fill=1)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 14)
    p.drawString(45, 727.5, "DATE")
    p.setFont("Helvetica", 14)
    
    p.setFillColor(colors.lightgrey)
    p.rect(140, 720, 415, 25, fill=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 14)
    p.drawString(145, 727.5, f"{formatted_date}")
    
    p.setFillColor(colors.lightgrey)
    p.rect(40, 684, 515, 25, fill=1, stroke=0)
    p.setFillColor(colors.black)
    p.setFont("Helvetica", 15)
    text_width = p.stringWidth("SUMMARY", "Helvetica", 15)
    x_center = (A4[0] - text_width) / 2
    p.drawString(x_center, 684+7.5, "SUMMARY")
    
    x = 40
    posisi_y = 654
    x2 = 195
    w = 25
    
    x_2 = 240
    x2_2 = 70
    
    x_3 = 315
    x2_3 = 165
    
    x_4 = 485
    x2_4 = 70
    
    # titles = [
    #     "Total Van Under Voltage", "Total Vbn Under Voltage", "Total Vcn Under Voltage", "Total Under Voltage",
    #     "Total Van Over Voltage", "Total Vbn Over Voltage", "Total Vcn Over Voltage", "Total Over Voltage",
    #     "Total Under Frequency", "Total Over Frequency",
    #     "Total High THDI Phase U", "Total High THDI Phase V", "Total High THDI Phase W", "Total High THDI",
    #     "Total High THDV Phase U", "Total High THDV Phase V", "Total High THDV Phase W", "Total High THDV",
    #     "Total Oil High Temp",
    #     "Total WTI Phase U High Temp", "Total WTI Phase V High Temp", "Total WTI Phase W High Temp", "Total WTI High Temp",
    #     "Total Low Power Factor Phase U", "Total Low Power Factor Phase V", "Total Low Power Factor Phase W", "Total Low Power Factor",
    #     "Total Over Current Phase U", "Total Over Current Phase V", "Total Over Current Phase W", "Total Over Current",
    #     "Total High Neutral Current",
    #     "Total Busbar Phase U High Temp", "Total Busbar Phase V High Temp", "Total Busbar Phase W High Temp", "Total Busbar High Temp",
    #     "Total High Pressure",
    #     "Total Van & Vbn Unbalance", "Total Van & Vcn Unbalance", "Total Vbn & Vcn Unbalance",
    #     "Total Vab & Vbc Unbalance", "Total Vab & Vca Unbalance", "Total Vbc & Vca Unbalance", "Total Unbalance"
    # ]
    
    titles = [
        "Total High THDI Phase U", "Total Van Under Voltage",
        "Total High THDI Phase V", "Total Vbn Under Voltage",
        "Total High THDI Phase W", "Total Vcn Under Voltage",
        "Total High THDI", "Total Under Voltage",
        "Total High THDV Phase U", "Total Van Over Voltage",
        "Total High THDV Phase V", "Total Vbn Over Voltage",
        "Total High THDV Phase W", "Total Vcn Over Voltage",
        "Total High THDV", "Total Over Voltage",
        "Total WTI Phase U High Temp", "Total Under Frequency",
        "Total WTI Phase V High Temp", "Total Over Frequency",
        "Total WTI Phase W High Temp", "Total Over Current Phase U",
        "Total WTI High Temp", "Total Over Current Phase V",
        "Total Oil High Temp", "Total Over Current Phase W",
        "Total Low Power Factor Phase U", "Total Over Current",
        "Total Low Power Factor Phase V", "Total High Neutral Current",
        "Total Low Power Factor Phase W", "Total Van & Vbn Unbalance",
        "Total Low Power Factor", "Total Van & Vcn Unbalance",
        "Total Busbar Phase U High Temp", "Total Vbn & Vcn Unbalance",
        "Total Busbar Phase V High Temp", "Total Vab & Vbc Unbalance",
        "Total Busbar Phase W High Temp", "Total Vab & Vca Unbalance",
        "Total Busbar High Temp", "Total Vbc & Vca Unbalance",
        "Total High Pressure", "Total Unbalance"
    ]
    
    values = [
        total_high_thd_i1, total_van_under_voltage,
        total_high_thd_i2, total_vbn_under_voltage,
        total_high_thd_i3, total_vcn_under_voltage,
        total_high_thd_i, total_under_voltage,
        total_high_thd_v1, total_van_over_voltage,
        total_high_thd_v2, total_vbn_over_voltage,
        total_high_thd_v3, total_vcn_over_voltage,
        total_high_thd_v, total_over_voltage,
        total_wti_hightemp1, total_under_freq,
        total_wti_hightemp2, total_over_freq,
        total_wti_hightemp3, total_over_currenta,
        total_wti_hightemp, total_over_currentb,
        total_oil_hightemp, total_over_currentc,
        total_low_pfa, total_over_current,
        total_low_pfb, total_high_ineutral,
        total_low_pfc, total_van_vbn_unbalance,
        total_low_pf, total_van_vcn_unbalance,
        total_bus_hightemp1, total_vbn_vcn_unbalance,
        total_bus_hightemp2, total_vab_vbc_unbalance,
        total_bus_hightemp3, total_vab_vca_unbalance,
        total_bus_hightemp, total_vbc_vca_unbalance,
        total_high_press, total_unbalance
    ]
    
    # values = [
    #     total_van_under_voltage, total_vbn_under_voltage, total_vcn_under_voltage, total_under_voltage,
    #     total_van_over_voltage, total_vbn_over_voltage, total_vcn_over_voltage, total_over_voltage,
    #     total_under_freq, total_over_freq,
    #     total_high_thd_i1, total_high_thd_i2, total_high_thd_i3, total_high_thd_i,
    #     total_high_thd_v1, total_high_thd_v2, total_high_thd_v3, total_high_thd_v,
    #     total_oil_hightemp,
    #     total_wti_hightemp1, total_wti_hightemp2, total_wti_hightemp3, total_wti_hightemp,
    #     total_low_pfa, total_low_pfb, total_low_pfc, total_low_pf,
    #     total_over_currenta, total_over_currentb, total_over_currentc, total_over_current,
    #     total_high_ineutral,
    #     total_bus_hightemp1, total_bus_hightemp2, total_bus_hightemp3, total_bus_hightemp,
    #     total_high_press,
    #     total_van_vbn_unbalance, total_van_vcn_unbalance, total_vbn_vcn_unbalance,
    #     total_vab_vbc_unbalance, total_vab_vca_unbalance, total_vbc_vca_unbalance, total_unbalance
    # ] 

    
    a = 0
    for i in range(0, 22):
        # Kotak Besar
        p.setFillColor(colors.lightgrey)
        p.rect(x, posisi_y, x2, w, fill=1, stroke=0)
        p.setFillColor(colors.black)
        p.setFont("Helvetica", 10)
        p.drawString(x+5, posisi_y+9, titles[a])

        # Kotak Kecil
        p.setFillColor(colors.lightgrey)
        p.rect(x_2, posisi_y, x2_2, w, fill=1, stroke=0)
        p.setFillColor(colors.black)
        p.setFont("Helvetica", 11)
        value_width = p.stringWidth(f"{values[a]}", "Helvetica", 10)
        p.drawString(x_2 + (x2_2 - value_width) / 2, posisi_y + 9, f"{values[a]}") 
        
        a += 1

        # Kotak Besar
        p.setFillColor(colors.lightgrey)
        p.rect(x_3, posisi_y, x2_3, w, fill=1, stroke=0)
        p.setFillColor(colors.black)
        p.setFont("Helvetica", 10)
        p.drawString(x_3+5, posisi_y+9, titles[a])

        # Kotak Kecil
        p.setFillColor(colors.lightgrey)
        p.rect(x_4, posisi_y, x2_4, w, fill=1, stroke=0)
        p.setFillColor(colors.black)
        p.setFont("Helvetica", 11)
        value_width = p.stringWidth(f"{values[a]}", "Helvetica", 10)
        p.drawString(x_4 + (x2_4 - value_width) / 2, posisi_y + 9, f"{values[a]}") 
        
        a += 1
        posisi_y -= 30
        
    
    
    # # Kotak Besar Pertama
    # p.setFillColor(colors.lightgrey)
    # p.rect(x, y, x2, w, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 11)
    # p.drawString(x+5, y+8.4, "Total Van & Vbn Unbalance")

    # # Kotak Kecil Pertama
    # p.setFillColor(colors.lightgrey)
    # p.rect(x_2, y, x2_2, w, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 11)
    # p.drawString(x_2+30, y+8.4, f"{total_van_vbn_unbalance}")

    # # Kotak Besar Kedua
    # p.setFillColor(colors.lightgrey)
    # p.rect(x_3, y, x2_3, w, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 11)
    # p.drawString(x_3+5, y+8.4, "Total Van & Vbn Unbalance")

    # # Kotak Kecil Kedua
    # p.setFillColor(colors.lightgrey)
    # p.rect(x_4, y, x2_4, w, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 11)
    # p.drawString(x_4+30, y+8.4  , f"{total_van_vbn_unbalance}")

    # # ----------------------------------------------------------
    
    # # Kotak Besar Pertama
    # p.setFillColor(colors.lightgrey)
    # p.rect(x, y-32, x2, w, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 11)
    # p.drawString(x+5, y+8.4-32, "Total Van & Vbn Unbalance")

    # # Kotak Kecil Pertama
    # p.setFillColor(colors.lightgrey)
    # p.rect(x_2, y-32, x2_2, w, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 11)
    # p.drawString(x_2+30, y+8.4-32, f"{total_van_vbn_unbalance}")

    # # Kotak Besar Kedua
    # p.setFillColor(colors.lightgrey)
    # p.rect(x_3, y-32, x2_3, w, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 11)
    # p.drawString(x_3+5, y+8.4-32, "Total Van & Vbn Unbalance")

    # # Kotak Kecil Kedua
    # p.setFillColor(colors.lightgrey)
    # p.rect(x_4, y-32, x2_4, w, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 11)
    # p.drawString(x_4+30, y+8.4-32, f"{total_van_vbn_unbalance}")
    
    

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 535, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 545, "Total Van & Vcn Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 535, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 545, f"{total_van_vcn_unbalance}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 500, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 510, "Total Vbn & Vcn Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 500, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 510, f"{total_vbn_vcn_unbalance}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 465, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 475, "Total Vab & Vbc Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 465, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 475, f"{total_vab_vbc_unbalance}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 430, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 440, "Total Vab & Vca Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 430, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 440, f"{total_vab_vca_unbalance}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 395, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 405, "Total Vbc & Vca Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 395, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 405, f"{total_vbc_vca_unbalance}")

    # p.setFillColor(colors.lightgrey)
    # p.rect(60, 360, 400, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(65, 370, "Total Unbalance")

    # p.setFillColor(colors.lightgrey)
    # p.rect(465, 360, 70, 30, fill=1, stroke=0)
    # p.setFillColor(colors.black)
    # p.setFont("Helvetica", 12)
    # p.drawString(495, 370, f"{total_unbalance}")
    
    p.showPage()
    p.setPageSize(landscape(A4))
    
    col_widths = [40, 35, 35, 33,
                33, 35, 35,
                47, 50, 31,
                35, 38, 40,
                35, 42, 217]
    
    chunk_size = 28
    chunks = [data[i:i + chunk_size] for i in range(0, len(data), chunk_size)]
    
    for chunk in chunks:
        table_data = [['timestamp', 'Under\nVoltage', 'Over\nVoltage', 'Under\nFrq',
                       'Over\nFrq', 'High\nTHD I', 'High\nTHD V',
                       'Oil\nHigh-temp', 'WTI\nHigh-temp', 'Low\nPF',
                       'Over\nCurrent', 'High\nI Neutral', 'Bus\nHigh-temp',
                       'High\nPress', 'Unbalance', 'Remarks']]
        
        for row in chunk:
            remarks = row.remarks.split(", ")
            formatted_remarks = ",\n".join([", ".join(remarks[i:i+9]) for i in range(0, len(remarks), 9)])

            table_data.append([row.timestamp, row.under_voltage, row.over_voltage, row.under_freq,
                               row.over_freq, row.high_thd_i, row.high_thd_v, row.oil_hightemp, row.wti_hightemp,
                               row.low_pf, row.over_current, row.high_ineutral, row.bus_hightemp,
                               row.high_press, row.unbalance, formatted_remarks])

        
        t = Table(table_data, colWidths=col_widths)
        
        style = [
                ('ROWHEIGHT', (0, 0), (-1, -1), 40),
                ('BACKGROUND', (0, 0), (-1, 0), '#4472C4'),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                ('FONT', (0, 0), (-1, -1), 'Helvetica', 7.5)]

        t.setStyle(style)
        
        cell_height = 20
        
        table_height = (len(chunk) + 1) * cell_height
        
        
        # y = (4.54545*len(chunk))+3.5455
        # y = 21.333*len(chunk)-153.33
        # y = 150
        # y = 230
        y = 4.444*len(chunk)+105.56
        
        table_y = 450 - table_height + (y)
        
        # print(t._rowHeights)
        t.wrapOn(p, 0, 0)
        t.drawOn(p, 30, table_y)
        p.showPage()

    download_name = f'{cust} Event Report'
    p.setTitle(download_name)
    
    p.save()
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name=f'Report.pdf')

if __name__ == '__main__':
    app.run(debug=True)
