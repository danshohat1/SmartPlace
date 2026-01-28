from pprint import pp
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify
import pandas as pd
import json
import os
import re
from logic import calculate_student_vectors, run_optimized_matching, run_full_optimization
import io

app = Flask(__name__)
app.secret_key = 'smartplace-secret-key-2026'  # נדרש עבור Flash messages
DB_FILE = 'db.json'

# --- ניהול נתונים ---
def load_db():
    default_structure = {"students": [], "units": {}, "saved_classes": {}}
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
                if 'saved_classes' not in data: data['saved_classes'] = {}
                return data
        except json.JSONDecodeError:
            return default_structure
    return default_structure

# --- Routes (עמודי האתר) ---
def save_db(data):
    with open(DB_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def read_uploaded_file(file_obj):
    """קורא קובץ Excel או CSV מהעלאה"""
    filename = file_obj.filename.lower()
    
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        return pd.read_excel(file_obj, engine='openpyxl' if filename.endswith('.xlsx') else None)
    elif filename.endswith('.csv'):
        return pd.read_csv(file_obj, encoding='utf-8-sig')
    else:
        raise ValueError("פורמט קובץ לא תומך. בחר Excel או CSV")

@app.route('/')
def index():
    data = load_db()
    unit_list = list(data['units'].keys())
    saved_classes = data.get('saved_classes', {})
    
    return render_template('index.html', 
                           units=unit_list, 
                           student_count=len(data['students']),
                           classes=saved_classes,
                           class_names=list(saved_classes.keys()))

@app.route('/upload', methods=['POST'])
def upload_files():
    # קבלת הקבצים מהטופס
    f_students = request.files.get('students_file')
    f_units = request.files.get('units_file')
    
    if f_students and f_units:
        try:
            # קריאת הנתונים מקובץ (Excel או CSV)
            df_s = read_uploaded_file(f_students)
            df_u = read_uploaded_file(f_units)
            
            # הרצת הלוגיקה (מה שכתבנו ב-logic.py)
            students_json = calculate_student_vectors(df_s, df_u)
            
            units_json = {}
            for _, row in df_u.iterrows():
                units_json[row['UnitName']] = {
                    "capacity": int(row['Capacity']),
                    "prefs": [],  # יתמלא בהמשך ידנית
                    "power": 1.0,  # ברירת מחדל
                    "sticky_power": False  # ברירת מחדל - לא נעול
                }
                
            # שמירה ל-db.json כדי שהמידע יישמר גם אם תסגור את השרת
            # חשוב: שומרים גם את הכיתות השמורות
            data = load_db()
            data['students'] = students_json
            data['units'] = units_json
            save_db(data)
            flash(f"✅ נטענו {len(students_json)} סטודנטים ו-{len(units_json)} יחידות בהצלחה!", 'success')
        except Exception as e:
            print(f"שגיאה בהעלאה: {e}")
            flash(f"❌ שגיאה בהעלאה: {str(e)}", 'danger')
        
    return redirect(url_for('index'))

@app.route('/upload_forms_excel', methods=['POST'])
def upload_forms_excel():
    """
    קבלת קובץ Excel מ-Microsoft Forms עם תשובות הסטודנטים.
    פורמט Forms: שורה ראשונה = כותרות השאלות
    כל שורה נוספת = תשובות סטודנט אחד
    """
    f = request.files.get('forms_file')
    
    if f:
        try:
            # קריאת הנתונים מ-Excel
            df = read_uploaded_file(f)
            
            # דלג על השורות הריקות בהתחלה (Forms לעיתים מוסיף כותרות מיוחדות)
            df = df.dropna(how='all')
            
            # אם העמודה הראשונה היא "שם מלא" או "응답자" או משהו דומה, זהו הקובץ הנכון
            print(f"עמודות בקובץ: {list(df.columns)}")
            
            # מצא עמודות שמכילות "שם" (כל הגרסאות אפשריות)
            name_col = None
            id_col = None
            
            for col in df.columns:
                col_str = str(col).lower()
                if 'שם' in col_str or 'name' in col_str:
                    name_col = col
                if 'תעודת זהות' in col_str or 'id' in col_str:
                    id_col = col
            
            if name_col is None:
                # אם לא נמצא, בחר את העמודה הראשונה
                name_col = df.columns[0]
            
            print(f"עמודת שם: {name_col}")
            
            processed_students = []
            
            # עיבוד כל שורה (סטודנט)
            for _, row in df.iterrows():
                name = str(row[name_col]).strip() if pd.notna(row[name_col]) else None
                
                if not name or name.lower() == 'nan' or name.lower() == 'סטודנט':
                    continue
                
                # אוספים את כל הדירוגים מכל העמודות (מלבד שם ותעודת זהות)
                ratings = []
                id_num = ""
                
                for col in df.columns:
                    col_str = str(col).lower()
                    
                    # דלג על עמודות שאינן דירוגים
                    if 'שם' in col_str or 'name' in col_str or 'timestamp' in col_str or 'זמן' in col_str:
                        continue
                    
                    if 'תעודת זהות' in col_str or 'id' in col_str:
                        if pd.notna(row[col]):
                            id_num = str(row[col]).strip()
                        continue
                    
                    # נסה לחלץ דירוג מהערך
                    if pd.notna(row[col]):
                        val = str(row[col]).strip()
                        if val and val.lower() != 'nan':
                            try:
                                rating = extract_rating(val)
                                ratings.append(rating)
                            except:
                                pass
                
                # אם יש לפחות דירוג אחד, שמור את הסטודנט
                if ratings or name:
                    processed_students.append({
                        "name": name,
                        "id": id_num,
                        "ratings": ratings
                    })
                    print(f"נטען סטודנט: {name}, דירוגים: {len(ratings)}")
            
            # שמירת תשובות הסטודנטים ב-db.json
            data = load_db()
            
            print(f"סך הכל סטודנטים חדשים: {len(processed_students)}")
            
            # מעדכן את רשימת הסטודנטים עם הדירוגים שלהם
            added_count = 0
            for student_data in processed_students:
                # מוודא שלא מוסיפים כפילויות
                existing = next((s for s in data['students'] if s.get('name', '').lower() == student_data['name'].lower()), None)
                if existing:
                    existing['ratings'] = student_data['ratings']
                    if student_data['id']:
                        existing['id'] = student_data['id']
                else:
                    data['students'].append(student_data)
                    added_count += 1
            
            save_db(data)
            print(f"נשמרו {added_count} סטודנטים חדשים")
            flash(f"✅ נטעינו בהצלחה {added_count} סטודנטים חדשים מהטופס!", 'success')
            
        except Exception as e:
            print(f"שגיאה בעיבוד Forms Excel: {e}")
            import traceback
            traceback.print_exc()
            flash(f"❌ שגיאה בעיבוד הקובץ: {str(e)}", 'danger')
        
    return redirect(url_for('index'))

def extract_rating(text):
    """חילוץ דירוג מטקסט"""
    if pd.isna(text) or str(text).strip() == "":
        return 3
    match = re.search(r'\d+', str(text))
    return int(match.group()) if match else 3

@app.route('/student/<student_name>')
def view_student_profile(student_name):
    """צפייה בפרטי סטודנט בודד - הדירוגים שלו ודירוג היחידות לו"""
    data = load_db()
    
    # חיפוש הסטודנט
    student = next((s for s in data['students'] if s.get('name') == student_name), None)
    
    if not student:
        flash(f"❌ סטודנט '{student_name}' לא נמצא", 'danger')
        return redirect(url_for('index'))
    
    # חישוב דירוגים מסודרים
    ratings_list = student.get('ratings', [])
    
    # דירוג העדפות היחידות
    prefs_list = student.get('prefs', [])
    
    # סטטיסטיקה
    avg_rating = sum(ratings_list) / len(ratings_list) if ratings_list else 0
    
    return render_template('student_profile.html',
                           student=student,
                           student_name=student_name,
                           ratings=ratings_list,
                           rating_count=len(ratings_list),
                           avg_rating=avg_rating,
                           prefs=prefs_list,
                           units_data=data['units'])

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
        
        # --- 1.5. עדכון Sticky Power (כוח ברזל) ---
        sticky_power = request.form.get('sticky_power') == 'on'
        data['units'][unit_name]['sticky_power'] = sticky_power
        
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
    current_sticky = data['units'][unit_name].get('sticky_power', False)
    
    return render_template('unit_ranking.html', 
                           unit=unit_name, 
                           students=all_student_names, 
                           current_ranks=current_ranks_display,
                           current_power=current_power,
                           current_sticky=current_sticky)

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
    """ריצה רגילה - עם ה-Power הנוכחי, מוצא רק Gamma אופטימלי"""
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
    
    # העברת Power ערכים הנוכחיים כמו calculated powers (לא שונו בריצה רגילה)
    calculated_powers = {u: data['units'][u]['power'] for u in data['units']}

    return render_template('results.html', 
                           matches=matches, 
                           reasons=reasons, 
                           units_data=data['units'],
                           calculated_powers=calculated_powers,
                           units_grouped=units_grouped,
                           unmatched=unmatched,
                           stats=stats,
                           message=f"שיבוץ הושלם (Gamma: {best_gamma})",
                           optimization_type="רגיל")

@app.route('/run_unified')
def run_unified():
    """ריצה מהירה עם אופטימיזציה מלאה - מוצא גם Gamma וגם Power אופטימליים"""
    data = load_db()
    
    # יצור דירוג אחיד - כל הסטודנטים בדרגה 1 בכל היחידות
    student_names = [s['name'] for s in data['students']]
    for unit_name in data['units']:
        data['units'][unit_name]['prefs'] = [student_names] if student_names else []
    
    # הרץ אופטימיזציה מלאה
    (matches, reasons), best_gamma, best_powers = run_full_optimization(
        data['students'], 
        data['units'], 
        iterations=200
    )
    
    # --- עדכון ה-Power המצוי ב-DB (רק ליחידות שאינן Sticky) ---
    for unit_name, new_power in best_powers.items():
        if not data['units'][unit_name].get('sticky_power', False):
            data['units'][unit_name]['power'] = new_power
    
    save_db(data)
    
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
                           calculated_powers=best_powers,
                           units_grouped=units_grouped,
                           unmatched=unmatched,
                           stats=stats,
                           message=f"🚀 ריצה מהירה עם אופטימיזציה מלאה הושלמה! (Gamma: {best_gamma})",
                           optimization_type="ריצה מהירה - אופטימיזציה")

@app.route('/run_full_optimization')
def run_full_opt():
    """
    ריצה אופטימלית מלאה - מוצא גם Gamma וגם Power אופטימליים.
    יחידות עם sticky_power=True ישמרו את ה-Power שלהן.
    """
    data = load_db()
    (matches, reasons), best_gamma, best_powers = run_full_optimization(
        data['students'], 
        data['units'], 
        iterations=200
    )
    
    # --- עדכון ה-Power המצוי ב-DB (רק ליחידות שאינן Sticky) ---
    for unit_name, new_power in best_powers.items():
        if not data['units'][unit_name].get('sticky_power', False):
            data['units'][unit_name]['power'] = new_power
    
    save_db(data)
    
    # --- הכנת נתונים לתצוגה ---
    units_grouped = {u: [] for u in data['units'].keys()}
    unmatched = []
    
    for student, unit in matches.items():
        if unit:
            if unit in units_grouped:
                units_grouped[unit].append(student)
        else:
            unmatched.append(student)
            
    stats = {
        "labels": list(data['units'].keys()),
        "assigned": [len(units_grouped[u]) for u in data['units']],
        "capacity": [data['units'][u]['capacity'] for u in data['units']]
    }

    return render_template('results.html', 
                           matches=matches, 
                           reasons=reasons, 
                           units_data=data['units'],
                           calculated_powers=best_powers,
                           units_grouped=units_grouped,
                           unmatched=unmatched,
                           stats=stats,
                           message=f"שיבוץ אופטימלי הושלם! (Gamma: {best_gamma})",
                           optimization_type="מלא")

@app.route('/run_class_optimized/<class_name>')
def run_class_optimized(class_name):
    """הרצה אופטימלית של כיתה שמורה"""
    data = load_db()
    saved_classes = data.get('saved_classes', {})
    
    if class_name not in saved_classes:
        flash(f"❌ הכיתה '{class_name}' לא נמצאה", 'danger')
        return redirect(url_for('classes_management'))
    
    class_data = saved_classes[class_name]
    students = class_data.get('students', [])
    units = class_data.get('units', {})
    
    try:
        (matches, reasons), best_gamma, best_powers = run_full_optimization(
            students, 
            units, 
            iterations=200
        )
        
        units_grouped = {u: [] for u in units.keys()}
        unmatched = []
        
        for student, unit in matches.items():
            if unit:
                if unit in units_grouped:
                    units_grouped[unit].append(student)
            else:
                unmatched.append(student)
        
        stats = {
            "labels": list(units.keys()),
            "assigned": [len(units_grouped[u]) for u in units],
            "capacity": [units[u]['capacity'] for u in units]
        }
        
        return render_template('results.html', 
                               matches=matches, 
                               reasons=reasons, 
                               units_data=units,
                               calculated_powers=best_powers,
                               units_grouped=units_grouped,
                               unmatched=unmatched,
                               stats=stats,
                               message=f"🎯 שיבוץ אופטימלי לכיתה '{class_name}' הושלם! (Gamma: {best_gamma})",
                               optimization_type="כיתה שמורה - אופטימיזציה")
    except Exception as e:
        print(f"שגיאה בהרצה: {e}")
        flash(f"❌ שגיאה בחישוב: {str(e)}", 'danger')
        return redirect(url_for('classes_management'))

@app.route('/units_management')
def units_management():
    """דף ניהול דירוג היחידות"""
    data = load_db()
    
    rankings = {}
    for unit_name, unit_data in data['units'].items():
        rankings[unit_name] = {
            'power': unit_data.get('power', 1.0),
            'sticky': unit_data.get('sticky_power', False),
            'ranks_count': len(unit_data.get('prefs', []))
        }
    
    return render_template('units_management.html', 
                           units=list(data['units'].keys()),
                           rankings=rankings)

@app.route('/upload_units_excel', methods=['POST'])
def upload_units_excel():
    """
    קבלת קובץ Excel עם דירוג כל היחידות.
    הפורמט המצופה:
    - עמודה ראשונה: שם היחידה
    - עמודות הבאות: שמות הסטודנטים
    - ערכים: דירוגים
    """
    f = request.files.get('units_excel')
    
    if f:
        try:
            df = read_uploaded_file(f)
            data = load_db()
            
            # קורא את העמודה הראשונה (שמות יחידות)
            first_col_name = df.columns[0]
            
            for _, row in df.iterrows():
                unit_name = str(row[first_col_name]).strip()
                
                if not unit_name or unit_name.lower() == 'nan':
                    continue
                
                # אם היחידה לא קיימת, תוסיף אותה
                if unit_name not in data['units']:
                    data['units'][unit_name] = {
                        "capacity": 5,  # ברירת מחדל
                        "prefs": [],
                        "power": 1.0,
                        "sticky_power": False
                    }
                
                # אוסף את הדירוגים מכל העמודות האחרות
                tier_map = {}
                for col_idx in range(1, len(df.columns)):
                    student_name = df.columns[col_idx]
                    rating_val = row.iloc[col_idx]
                    
                    if pd.notna(rating_val) and str(rating_val).strip() != '':
                        try:
                            rank = int(rating_val)
                            if rank not in tier_map:
                                tier_map[rank] = []
                            tier_map[rank].append(student_name)
                        except:
                            pass
                
                # מיון הקבוצות לפי דירוג
                sorted_keys = sorted(tier_map.keys())
                final_prefs = [tier_map[k] for k in sorted_keys]
                
                data['units'][unit_name]['prefs'] = final_prefs
            
            save_db(data)
        except Exception as e:
            print(f"שגיאה בעיבוד Units Excel: {e}")
    
    return redirect(url_for('units_management'))

@app.route('/download_units_template')
def download_units_template():
    """הורדה של תבנית Excel לדירוג יחידות"""
    data = load_db()
    
    # יצירת דאטא לדוגמה
    template_data = {
        "שם יחידה": list(data['units'].keys()) if data['units'] else ["יחידה 1", "יחידה 2"]
    }
    
    # הוספת שמות סטודנטים כעמודות
    student_names = [s.get('name', f'סטודנט {i}') for i, s in enumerate(data['students'])]
    for name in student_names[:10]:  # עד 10 סטודנטים בתבנית
        template_data[name] = [1, 2] * 5  # דוגמה לדירוגים
    
    df = pd.DataFrame(template_data)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='דירוג יחידות')
        
        worksheet = writer.sheets['דירוג יחידות']
        for column_cells in worksheet.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2
    
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='units_ranking_template.xlsx'
    )

@app.route('/download_students_sample')
def download_students_sample():
    """הורדה של קובץ דוגמה לסטודנטים"""
    sample_data = {
        'שם': ['דן כהן', 'ירון לוי', 'טל בן דור', 'שרה גולן', 'ניסן ברקוביץ'],
        'שאלה 1': [3, 2, 1, 3, 2],
        'שאלה 2': [1, 3, 2, 2, 1],
        'שאלה 3': [2, 1, 3, 1, 3],
    }
    
    df = pd.DataFrame(sample_data)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='דירוגים')
        
        worksheet = writer.sheets['דירוגים']
        for column_cells in worksheet.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2
    
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='sample_students.xlsx'
    )

@app.route('/download_units_sample')
def download_units_sample():
    """הורדה של קובץ דוגמה ליחידות"""
    sample_data = {
        'UnitName': ['8200', 'מטכ״ל', 'יחידת תוכנה', 'קשר חטיבתי'],
        'Capacity': [10, 2, 15, 5],
        'Q1': [1, 5, 1, 4],
        'Q2': [4, 5, 5, 2],
        'Q3': [5, 5, 2, 3],
        'Q4': [1, 5, 1, 4],
        'Q5': [3, 5, 2, 4],
        'Q6': [2, 1, 5, 2],
        'Q7': [2, 1, 4, 2],
        'Q8': [3, 1, 5, 2],
        'Q9': [4, 1, 5, 2],
    }
    
    df = pd.DataFrame(sample_data)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='יחידות')
        
        worksheet = writer.sheets['יחידות']
        for column_cells in worksheet.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2
    
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='sample_units.xlsx'
    )


@app.route('/classes')
def classes_management():
    """דף ניהול כיתות שמורות"""
    data = load_db()
    saved_classes = data.get('saved_classes', {})
    
    return render_template('classes.html', 
                           classes=saved_classes,
                           class_names=list(saved_classes.keys()))

@app.route('/save_class', methods=['POST'])
def save_class():
    """שמירת כיתה חדשה"""
    class_name = request.form.get('class_name', '').strip()
    class_description = request.form.get('class_description', '').strip()
    
    if not class_name:
        return redirect(url_for('classes_management'))
    
    data = load_db()
    
    # וידוא שקיים מפתח saved_classes
    if 'saved_classes' not in data:
        data['saved_classes'] = {}
    
    # שמירת הכיתה עם המצב הנוכחי
    # חשוב: אנחנו משמרים עותק מהנתונים, לא הפניה
    current_students = json.loads(json.dumps(data.get('students', [])))
    current_units = json.loads(json.dumps(data.get('units', {})))
    
    data['saved_classes'][class_name] = {
        'students': current_students,
        'units': current_units,
        'description': class_description,
        'created_date': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')
    }
    
    save_db(data)
    return redirect(url_for('classes_management'))

@app.route('/load_class/<class_name>')
def load_class(class_name):
    """טעינת כיתה שמורה וביצוע חישוב"""
    data = load_db()
    saved_classes = data.get('saved_classes', {})
    
    if class_name not in saved_classes:
        return redirect(url_for('classes_management'))
    
    class_data = saved_classes[class_name]
    
    # עתיד לטמפורריות - לא שומרים בחזרה ל-db
    students = class_data.get('students', [])
    units = class_data.get('units', {})
    
    # ריצת חישוב אופטימלי
    try:
        (matches, reasons), best_gamma = run_optimized_matching(students, units)
        
        units_grouped = {u: [] for u in units.keys()}
        unmatched = []
        
        for student, unit in matches.items():
            if unit:
                if unit in units_grouped:
                    units_grouped[unit].append(student)
            else:
                unmatched.append(student)
        
        stats = {
            "labels": list(units.keys()),
            "assigned": [len(units_grouped[u]) for u in units],
            "capacity": [units[u]['capacity'] for u in units]
        }
        
        return render_template('results.html', 
                               matches=matches, 
                               reasons=reasons, 
                               units_data=units,
                               units_grouped=units_grouped,
                               unmatched=unmatched,
                               stats=stats,
                               message=f"שיבוץ לכיתה '{class_name}' הושלם! (Gamma: {best_gamma})",
                               optimization_type="כיתה שמורה")
    except Exception as e:
        print(f"שגיאה בטעינת כיתה: {e}")
        return redirect(url_for('classes_management'))

@app.route('/edit_class/<class_name>')
def edit_class(class_name):
    """עדכון כיתה שמורה"""
    data = load_db()
    saved_classes = data.get('saved_classes', {})
    
    if class_name not in saved_classes:
        return redirect(url_for('classes_management'))
    
    class_data = saved_classes[class_name]
    
    return render_template('edit_class.html',
                           class_name=class_name,
                           class_data=class_data,
                           student_count=len(class_data.get('students', [])),
                           unit_count=len(class_data.get('units', {})))

@app.route('/update_class', methods=['POST'])
def update_class():
    """עדכון דירוגים של יחידות בכיתה שמורה"""
    try:
        request_data = request.get_json()
        class_name = request_data.get('class_name')
        updated_units = request_data.get('units', {})
        
        if not class_name:
            return jsonify({'success': False, 'error': 'שם הכיתה חסר'}), 400
        
        data = load_db()
        saved_classes = data.get('saved_classes', {})
        
        if class_name not in saved_classes:
            print(f"ERROR: Class {class_name} not found in saved_classes")
            print(f"Available classes: {list(saved_classes.keys())}")
            return jsonify({'success': False, 'error': 'כיתה לא נמצאה'}), 404
        
        print(f"Updating class: {class_name}")
        print(f"Units count: {len(updated_units)}")
        
        # עדכון היחידות בכיתה השמורה
        saved_classes[class_name]['units'] = updated_units
        data['saved_classes'] = saved_classes
        
        # שמירה לDB
        save_db(data)
        print(f"Class {class_name} updated successfully")
        
        return jsonify({'success': True, 'message': f'כיתה {class_name} עודכנה בהצלחה'})
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error updating class: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/delete_class/<class_name>', methods=['POST'])
def delete_class(class_name):
    """מחיקת כיתה שמורה"""
    data = load_db()
    saved_classes = data.get('saved_classes', {})
    
    if class_name in saved_classes:
        del data['saved_classes'][class_name]
        save_db(data)
    
    return redirect(url_for('classes_management'))

if __name__ == '__main__':
    app.run(debug=True, port=5001)