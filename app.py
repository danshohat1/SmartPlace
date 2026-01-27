from pprint import pp
from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import json
import os
from SmartPlace.logic import calculate_student_vectors, run_optimized_matching
import io
from flask import send_file

app = Flask(__name__)
DB_FILE = 'db.json'

# --- ניהול נתונים ---
def load_db():
    default_structure = {"students": [], "units": {}}
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if not content:
                    return default_structure
                data = json.loads(content)
                # וידוא שהמפתחות קיימים גם אם הקובץ חלקי
                if 'units' not in data: data['units'] = {}
                if 'students' not in data: data['students'] = []
                return data
        except json.JSONDecodeError:
            return default_structure
    return default_structure

# --- Routes (עמודי האתר) ---
def save_db(data):
    with open(DB_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

@app.route('/')
def index():
    data = load_db()
    unit_list = list(data['units'].keys())
    return render_template('index.html', units=unit_list, student_count=len(data['students']))

@app.route('/upload', methods=['POST'])
def upload_files():
    # קבלת הקבצים מהטופס
    f_students = request.files.get('students_file')
    f_units = request.files.get('units_file')
    
    if f_students and f_units:
        # קריאת הנתונים ישירות מהקובץ שהועלה
        df_s = pd.read_csv(f_students, encoding='utf-8-sig')
        df_u = pd.read_csv(f_units, encoding='utf-8-sig')
        
        # הרצת הלוגיקה (מה שכתבנו ב-logic.py)
        students_json = calculate_student_vectors(df_s, df_u)
        
        units_json = {}
        for _, row in df_u.iterrows():
            units_json[row['UnitName']] = {
                "capacity": int(row['Capacity']),
                "prefs": []  # יתמלא בהמשך ידנית
            }
            
        # שמירה ל-db.json כדי שהמידע יישמר גם אם תסגור את השרת
        save_db({"students": students_json, "units": units_json})
        
    return redirect(url_for('index'))

@app.route('/rank/<unit_name>', methods=['GET', 'POST'])
def rank_unit(unit_name):
    data = load_db()
    if unit_name not in data['units']:
        return redirect(url_for('index'))
    
    # רשימת כל השמות של הסטודנטים במערכת
    all_student_names = [s['name'] for s in data['students']]
    
    if request.method == 'POST':
        # --- 1. עדכון כוח היחידה (Power) ---
        new_power = float(request.form.get('unit_power', 1.0))
        data['units'][unit_name]['power'] = new_power
        
        # --- 2. עיבוד דירוגי הסטודנטים ---
        # אנחנו אוספים את כל הדירוגים מהטופס (רק מה שאינו ריק)
        ranks = {}
        for name in all_student_names:
            val = request.form.get(f'rank_{name}')
            if val and val.strip():
                ranks[name] = int(val)
        
        # הפיכת הדירוגים למבנה של Tiers (קבוצות)
        # דוגמה: { "דן": 1, "יוסי": 1, "ערן": 2 } -> [ ["דן", "יוסי"], ["ערן"] ]
        tier_map = {}
        for name, rank_val in ranks.items():
            if rank_val not in tier_map:
                tier_map[rank_val] = []
            tier_map[rank_val].append(name)
        
        # מיון הקבוצות לפי מספר הדירוג (מהקטן לגדול)
        sorted_keys = sorted(tier_map.keys())
        final_prefs = [tier_map[k] for k in sorted_keys]
        
        data['units'][unit_name]['prefs'] = final_prefs
        save_db(data)
        return redirect(url_for('index'))

    # למטרת תצוגה ב-GET: ננסה להבין מה הדירוג הנוכחי של כל סטודנט כדי להציג אותו בתיבות
    current_prefs = data['units'][unit_name].get('prefs', [])
    current_ranks_display = {}
    for i, tier in enumerate(current_prefs):
        names = tier if isinstance(tier, list) else [tier]
        for n in names:
            current_ranks_display[n] = i + 1
            
    current_power = data['units'][unit_name].get('power', 1.0)
    
    return render_template('unit_ranking.html', 
                           unit=unit_name, 
                           students=all_student_names, 
                           current_ranks=current_ranks_display,
                           current_power=current_power)

@app.route('/download_excel')
def download_excel():
    data = load_db()
    # נריץ חישוב מהיר כדי לקבל את המצב העדכני (או שנשמור ב-session, כרגע נריץ שוב)
    (matches, reasons), _ = run_optimized_matching(data['students'], data['units'])
    
    # בניית הדאטה-פריים לאקסל
    export_data = []
    for student_name, unit_name in matches.items():
        export_data.append({
            "שם הסטודנט": student_name,
            "יחידה משובצת": unit_name if unit_name else "לא שובץ",
            "הסבר לשיבוץ": reasons.get(student_name, ""),
            "כוח היחידה": data['units'].get(unit_name, {}).get('power', 0) if unit_name else 0
        })
    
    df = pd.DataFrame(export_data)
    
    # יצירת קובץ בזיכרון (ללא שמירה לדיסק)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='תוצאות שיבוץ')
        
        # התאמת רוחב עמודות (קוסמטיקה לאקסל)
        worksheet = writer.sheets['תוצאות שיבוץ']
        for column_cells in worksheet.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2

    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='placement_results.xlsx'
    )

@app.route('/run')
def run_matching():
    data = load_db()
    (matches, reasons), best_gamma = run_optimized_matching(data['students'], data['units'])
    
    # --- הכנת נתונים לתצוגה לפי יחידות (Grouping) ---
    units_grouped = {u: [] for u in data['units'].keys()}
    unmatched = []
    
    for student, unit in matches.items():
        if unit:
            if unit in units_grouped:
                units_grouped[unit].append(student)
        else:
            unmatched.append(student)
            
    # חישוב תפוסה לגרפים
    stats = {
        "labels": list(data['units'].keys()),
        "assigned": [len(units_grouped[u]) for u in data['units']],
        "capacity": [data['units'][u]['capacity'] for u in data['units']]
    }

    return render_template('results.html', 
                           matches=matches, 
                           reasons=reasons, 
                           units_data=data['units'],
                           units_grouped=units_grouped, # העברת המידע המקובץ
                           unmatched=unmatched,
                           stats=stats, # העברת מידע לגרפים
                           message=f"שיבוץ הושלם (Gamma: {best_gamma})")
if __name__ == '__main__':
    app.run(debug=True, port=5001)