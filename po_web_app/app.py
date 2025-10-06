from flask import Flask, render_template, request, redirect
import pandas as pd
import os
import datetime

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def generate_po_number(index):
    return f"{index}/69"

#EXCEL_PATH = os.path.join(UPLOAD_FOLDER, 'latest.xlsx')  # ใช้ชื่อไฟล์คงที่

@app.route('/', methods=['GET', 'POST'])
def index():
    po_data = None
    latest_path = None

    if request.method == 'POST':
        file = request.files['excel_file']
        

        # ✅ ตั้งชื่อไฟล์ใหม่ทุกครั้ง
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"po_{timestamp}.xlsx"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        
        latest_path = filepath  # ใช้ไฟล์ล่าสุดแสดงหน้าแรก

        df = pd.read_excel(filepath)

        for col in ['PO Number', 'สถานะการจัดส่ง','ค้างส่ง']:
            if col not in df.columns:
                df[col] = ''

        # ✅ หาเลข PO ล่าสุดจากทุกไฟล์
        existing_numbers = []
        for fname in os.listdir(UPLOAD_FOLDER):
            if fname.endswith('.xlsx'):
                try:
                    temp_df = pd.read_excel(os.path.join(UPLOAD_FOLDER, fname))
                    pos = temp_df['PO Number'].dropna().astype(str)
                    nums = [int(po.split('/')[0]) for po in pos if '/69' in po]
                    existing_numbers.extend(nums)
                except:
                    continue

        start_index = max(existing_numbers) + 1 if existing_numbers else 1
        
        previous_company = None
        current_po = None

        for i in range(len(df)):
            company = str(df.loc[i, 'ชื่อบริษัท']).strip().lower() if 'ชื่อบริษัท' in df.columns else 'ไม่ระบุ'

            # ถ้าเป็น GPO ให้ขึ้นว่า "งวดยา GPO"
            if company == 'gpo':
                df.loc[i, 'PO Number'] = 'งวดยา GPO'
            elif not df.loc[i, 'PO Number']:
                if company == previous_company:
                    df.loc[i, 'PO Number'] = current_po
                else:
                    current_po = generate_po_number(start_index)
                    df.loc[i, 'PO Number'] = current_po
                    start_index += 1
                previous_company = company

            # สถานะการจัดส่ง
            if not df.loc[i, 'สถานะการจัดส่ง']:
                df.loc[i, 'สถานะการจัดส่ง'] = 'ยังไม่ส่ง'

            if not df.loc[i, 'สถานะการจัดส่ง']:
                df.loc[i, 'สถานะการจัดส่ง'] = 'ยังไม่ส่ง'

        df.to_excel(filepath, index=False)

    
    # ✅ โหลดไฟล์ล่าสุดจาก uploads
    files = sorted([f for f in os.listdir(UPLOAD_FOLDER) if f.endswith('.xlsx')], reverse=True)
    if files:
        latest_path = os.path.join(UPLOAD_FOLDER, files[0])
        df = pd.read_excel(latest_path)

        for col in ['PO Number', 'ชื่อเวชภัณฑ์', 'ปริมาณจัดซื้อ', 'หน่วยนับ', 'ชื่อบริษัท']:
            if col not in df.columns:
                df[col] = ''

        selected_columns = ['PO Number', 'ชื่อเวชภัณฑ์', 'ปริมาณจัดซื้อ', 'หน่วยนับ', 'ชื่อบริษัท']
        po_data = df[selected_columns].to_dict(orient='records')

    return render_template('index.html', po_data=po_data)

@app.route('/update_status', methods=['POST'])
def update_status():
    po_number = request.form.get('po_number')
    new_status = request.form.get('new_status')
    pending_qty = request.form.get('pending_qty')

    if os.path.exists(EXCEL_PATH):
        df = pd.read_excel(EXCEL_PATH)

        # อัปเดตสถานะ
        df.loc[df['PO Number'] == po_number, 'สถานะการจัดส่ง'] = new_status

        # เพิ่มคอลัมน์ "ค้างส่ง" ถ้ายังไม่มี
        if 'ค้างส่ง' not in df.columns:
            df['ค้างส่ง'] = ''

        # อัปเดตจำนวนค้างส่ง
        if pending_qty is not None:
            df.loc[df['PO Number'] == po_number, 'ค้างส่ง'] = pending_qty

        df.to_excel(EXCEL_PATH, index=False)

    return redirect('/')


@app.route('/history')
def history():
    all_data = []
    all_items = set()
    all_companies = set()

    for fname in sorted(os.listdir(UPLOAD_FOLDER), reverse=True):
        if fname.endswith('.xlsx'):
            path = os.path.join(UPLOAD_FOLDER, fname)
            try:
                df = pd.read_excel(path)

                for col in ['PO Number', 'ชื่อเวชภัณฑ์', 'ปริมาณจัดซื้อ', 'หน่วยนับ', 'ชื่อบริษัท', 'สถานะการจัดส่ง', 'ค้างส่ง']:
                    if col not in df.columns:
                        df[col] = ''

                # เก็บชื่อเวชภัณฑ์และบริษัท
                all_items.update(df['ชื่อเวชภัณฑ์'].dropna().astype(str).str.strip())
                all_companies.update(df['ชื่อบริษัท'].dropna().astype(str).str.strip())

                records = df[['PO Number', 'ชื่อเวชภัณฑ์', 'ปริมาณจัดซื้อ', 'หน่วยนับ', 'ชื่อบริษัท', 'สถานะการจัดส่ง', 'ค้างส่ง']].to_dict(orient='records')
                all_data.append({'filename': fname, 'records': records})
            except:
                continue

    return render_template('history.html', all_data=all_data,
                           item_options=sorted(all_items),
                           company_options=sorted(all_companies))


@app.route('/update_history_status', methods=['POST'])
def update_history_status():
    filename = request.form.get('filename')
    row_id = int(request.form.get('row_id')) 
    new_status = request.form.get('new_status')
    pending_qty = request.form.get('pending_qty')

    path = os.path.join(UPLOAD_FOLDER, filename)
    if os.path.exists(path):
        df = pd.read_excel(path)

        if 'สถานะการจัดส่ง' not in df.columns:
            df['สถานะการจัดส่ง'] = ''
        if 'ค้างส่ง' not in df.columns:
            df['ค้างส่ง'] = ''

        df.at[row_id, 'สถานะการจัดส่ง'] = new_status
        try:
            pending_qty = int(pending_qty)
            df.at[row_id, 'ค้างส่ง'] = pending_qty
        except (ValueError, TypeError):
            pass # ข้ามถ้าไม่ใช่ตัวเลข


        df.to_excel(path, index=False)

    return redirect('/history')

@app.route('/delete_history_row', methods=['POST'])
def delete_history_row():
    filename = request.form.get('filename')
    po_number = request.form.get('po_number')

    path = os.path.join(UPLOAD_FOLDER, filename)
    if os.path.exists(path):
        df = pd.read_excel(path)

        # แปลง PO Number เป็น string และลบช่องว่าง
        df['PO Number'] = df['PO Number'].astype(str).str.strip()
        po_number = str(po_number).strip()

        # ลบแถวที่ตรงกับ PO Number
        df = df[df['PO Number'] != po_number]

        df.to_excel(path, index=False)

    return redirect('/history')

@app.route('/add_entry', methods=['POST'])
def add_entry():
    filename = request.form.get('filename')
    path = os.path.join(UPLOAD_FOLDER, filename)

    if os.path.exists(path):
        df = pd.read_excel(path)

        # สร้าง PO ใหม่
        existing_numbers = df['PO Number'].dropna().astype(str)
        nums = [int(po.split('/')[0]) for po in existing_numbers if '/69' in po]
        start_index = max(nums) + 1 if nums else 1
        new_po = f"{start_index}/69"

        # รับข้อมูลจากฟอร์ม
        new_row = {
            'PO Number': new_po,
            'ชื่อเวชภัณฑ์': request.form.get('ชื่อเวชภัณฑ์'),
            'ปริมาณจัดซื้อ': request.form.get('ปริมาณจัดซื้อ'),
            'หน่วยนับ': request.form.get('หน่วยนับ'),
            'ชื่อบริษัท': request.form.get('ชื่อบริษัท'),
            'สถานะการจัดส่ง': 'ยังไม่ส่ง',
            'ค้างส่ง': ''
        }

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(path, index=False)

    return redirect('/history')





@app.route('/add_manual_entry', methods=['POST'])
def add_manual_entry():
    # สร้างชื่อไฟล์ใหม่
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"manual_{timestamp}.xlsx"
    path = os.path.join(UPLOAD_FOLDER, filename)

    # สร้าง PO ใหม่
    existing_numbers = []
    for fname in os.listdir(UPLOAD_FOLDER):
        if fname.endswith('.xlsx'):
            try:
                temp_df = pd.read_excel(os.path.join(UPLOAD_FOLDER, fname))
                pos = temp_df['PO Number'].dropna().astype(str)
                nums = [int(po.split('/')[0]) for po in pos if '/69' in po]
                existing_numbers.extend(nums)
            except:
                continue
    start_index = max(existing_numbers) + 1 if existing_numbers else 1
    new_po = f"{start_index}/69"

    # สร้างแถวใหม่
    new_row = {
        'PO Number': new_po,
        'ชื่อเวชภัณฑ์': request.form.get('ชื่อเวชภัณฑ์'),
        'ปริมาณจัดซื้อ': request.form.get('ปริมาณจัดซื้อ'),
        'หน่วยนับ': request.form.get('หน่วยนับ'),
        'ชื่อบริษัท': request.form.get('ชื่อบริษัท'),
        'สถานะการจัดส่ง': 'ยังไม่ส่ง',
        'ค้างส่ง': ''
    }

    df = pd.DataFrame([new_row])
    df.to_excel(path, index=False)

    return redirect('/history')





if __name__ == '__main__':
    app.run(debug=True)