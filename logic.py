import pandas as pd
import numpy as np
import re
from dataclasses import dataclass, field
from typing import List, Union, Dict, Optional, Tuple

# --- 1. מודלים (Models) עם כל המתודות הנדרשות ---

@dataclass
class Student:
    name: str
    preferences: List[str]
    voice: float = 1.0
    match: Optional[str] = None

@dataclass
class University:
    name: str
    capacity: int
    preferences: List[Union[str, List[str]]]
    power: float = 1.0
    accepted: List[str] = field(default_factory=list)
    preferences_flat: List[str] = field(init=False)
    preference_pointer: int = field(default=0, init=False)

    def __post_init__(self):
        self.preferences_flat = []
        for tier in self.preferences:
            if isinstance(tier, list):
                self.preferences_flat.extend(tier)
            else:
                self.preferences_flat.append(tier)

    def has_free_slot(self) -> bool:
        return len(self.accepted) < self.capacity

    def next_candidate(self) -> Union[str, None]:
        """מחזירה את המועמד הבא בתור שהיחידה טרם בדקה"""
        while self.preference_pointer < len(self.preferences_flat):
            candidate = self.preferences_flat[self.preference_pointer]
            self.preference_pointer += 1
            return candidate
        return None

# --- 2. לוגיקה של ניתוח סקרים (Vectors) ---

def extract_rating(text):
    if pd.isna(text) or str(text).strip() == "": return 3
    match = re.search(r'\d+', str(text))
    return int(match.group()) if match else 3

def calculate_student_vectors(df_students, df_units):
    processed_students = []
    df_students = df_students.dropna(subset=['שם מלא'])
    q_cols = [c for c in df_students.columns if '?' in c and 'הבהרה' not in c]
    
    for _, row in df_students.iterrows():
        name = str(row['שם מלא']).strip()
        if not name or name.lower() == 'nan': continue
            
        s_vec = np.array([extract_rating(row[c]) for c in q_cols])
        dists = []
        for _, unit in df_units.iterrows():
            u_vec = np.array(unit.iloc[2:2+len(q_cols)].values, dtype=float)
            dist = np.linalg.norm(s_vec - u_vec)
            dists.append((unit['UnitName'], dist))
            
        sorted_units = [d[0] for d in sorted(dists, key=lambda x: x[1])]
        processed_students.append({"name": name, "prefs": sorted_units, "voice": 1.0})
        
    return processed_students

# --- 3. האלגוריתם המלא (weighted_gale_shapley) ששלחת ---

def get_rank(student: Student, university_name: str) -> int:
    for i, tier in enumerate(student.preferences):
        if university_name == tier or (isinstance(tier, list) and university_name in tier):
            return i
    return len(student.preferences)

def weighted_gale_shapley(students, universities, gamma=1.0):
    """
    מימוש אלגוריתם גייל-שפלי עם משקולות (Power) והסברים בעברית.
    יחידות (Universities) הן אלו שמציעות לסטודנטים.
    """
    n = len(universities)
    matches = {}
    reasons = {s: "" for s in students}

    def components(s, u):
        """חישוב רכיבי הניקוד: עדיפות הסטודנט מול כוח היחידה"""
        r = get_rank(s, u.name)
        # ככל שהדירוג (r) נמוך יותר (קרוב ל-0), הציון גבוה יותר
        stu_comp = s.voice * (n - r)
        uni_comp = gamma * u.power
        return r, stu_comp, uni_comp, stu_comp + uni_comp

    # יחידות עם מקום פנוי מתחילות להציע
    # מיון ראשוני לפי כוח (Power) כדי שיחידות חזקות יציעו קודם
    free_unis = sorted(
        [u for u in universities.values() if u.has_free_slot()],
        key=lambda x: x.power,
        reverse=True
    )

    while free_unis:
        uni = free_unis.pop(0)
        cand_name = uni.next_candidate()
        
        if not cand_name or cand_name not in students:
            continue
        
        stu = students[cand_name]
        r_new, stu_comp_new, uni_comp_new, total_new = components(stu, uni)

        # תרחיש 1: הסטודנט אינו משובץ כרגע
        if stu.match is None:
            stu.match = uni.name
            uni.accepted.append(cand_name)
            
            if r_new == 0:
                reasons[cand_name] = f"שובץ ל{uni.name} כי זו העדיפות הראשונה שלו."
            else:
                # חישוב הסף שבו הכוח של היחידה ניצח את העדיפות הראשונה
                # (voice * (n-0) + gamma * 1) vs (voice * (n-r_new) + gamma * Power)
                threshold = (stu.voice * r_new) / gamma + 1.0
                reasons[cand_name] = (
                    f"שובץ ל{uni.name} (עדיפות {r_new+1}) למרות שהיו לו העדפות גבוהות יותר. "
                    f"הסיבה: כוח היחידה ({uni.power}) גבר על העדפותיו האישיות. "
                    f"נקודת המפנה (Power Threshold) עבורו הייתה {threshold:.1f}."
                )
        
        # תרחיש 2: הסטודנט כבר משובץ ליחידה אחרת - בודקים האם להחליף
        else:
            current_uni = universities[stu.match]
            r_old, stu_comp_old, uni_comp_old, total_old = components(stu, current_uni)
            
            if total_new > total_old:
                # הסבר על ההחלפה
                pref_improved = r_new < r_old
                power_diff = uni_comp_new - uni_comp_old
                
                if pref_improved:
                    reasons[cand_name] = (
                        f"הועבר מ{current_uni.name} ל{uni.name} מכיוון שזו עדיפות גבוהה יותר "
                        f"(ממקום {r_old+1} למקום {r_new+1})."
                    )
                else:
                    reasons[cand_name] = (
                        f"הועבר מ{current_uni.name} ל{uni.name} למרות שהעדיפות נמוכה יותר, "
                        f"בשל פער כוח משמעותי לטובת היחידה החדשה."
                    )
                
                # ביצוע ההחלפה בפועל
                current_uni.accepted.remove(cand_name)
                uni.accepted.append(cand_name)
                stu.match = uni.name
                
                # היחידה שאיבדה סטודנט חוזרת לרשימת המציעים אם התפנה לה מקום
                if current_uni.has_free_slot():
                    free_unis.append(current_uni)
            else:
                # הסטודנט נשאר בשיבוץ הקיים
                pass

        # אם ליחידה עדיין יש מקום והיא לא סיימה את רשימת המועמדים שלה, היא ממשיכה להציע
        if uni.has_free_slot() and uni.preference_pointer < len(uni.preferences_flat):
            if uni not in free_unis:
                free_unis.append(uni)

    # סיכום תוצאות
    final_matches = {name: s.match for name, s in students.items()}
    
    # טיפול במי שלא שובץ
    for name, s in students.items():
        if s.match is None:
            reasons[name] = "לא נמצא שיבוץ; היחידות שהציעו לא היו בעלות משקל מספיק מול העדפות הסטודנט."

    return final_matches, reasons
# --- 4. אופטימיזציה (כמו ב-CLI) ---

def boost_voice_by_demand(students, universities, alpha=1.0):
    counts = {name: 0 for name in students}
    for uni in universities.values():
        for cand in uni.preferences_flat:
            if cand in counts: counts[cand] += 1
    for name, s in students.items():
        s.voice += alpha * counts[name]

def run_optimized_matching(students_data, units_data):
    """מריץ אופטימיזציה למציאת Gamma אידיאלי"""
    def reset_data():
        s = {sd['name']: Student(sd['name'], sd['prefs'], sd['voice']) for sd in students_data}
        u = {name: University(name, ud['capacity'], ud['prefs'], ud.get('power', 1.0)) 
             for name, ud in units_data.items()}
        return s, u

    best_gamma = 1.0
    fewest_unmatched = float('inf')
    final_results = None

    for g in np.arange(0.5, 3.0, 0.5):
        s, u = reset_data()
        boost_voice_by_demand(s, u) # כמו ב-CLI
        m, r = weighted_gale_shapley(s, u, gamma=g)
        unmatched = sum(1 for v in m.values() if v is None)
        if unmatched < fewest_unmatched:
            fewest_unmatched = unmatched
            best_gamma = g
            final_results = (m, r)

    return final_results, best_gamma

def optimize_run(students_data, units_data, iterations=100):
    best_matches = None
    best_reasons = None
    best_unmatched_count = float('inf')
    best_powers = {}

    # יצירת אובייקטים בסיסיים
    for _ in range(iterations):
        current_units_data = copy.deepcopy(units_data)
        
        # אופטימיזציה: שינוי Power ליחידות שאינן Sticky
        for u_name, u_info in current_units_data.items():
            if not u_info.get('sticky', False):
                # הגרלת כוח חדש בטווח הגיוני (למשל 0.5 עד 15.0)
                u_info['power'] = round(random.uniform(0.5, 15.0), 1)
        
        # הרצת האלגוריתם עם הנתונים החדשים
        # (כאן קוראים ל-weighted_gale_shapley המוכר)
        students_obj = {sd['name']: Student(sd['name'], sd['prefs'], sd['voice']) for sd in students_data}
        units_obj = {name: University(name, ui['capacity'], ui['prefs'], ui['power']) 
                     for name, ui in current_units_data.items()}
        
        matches, reasons = weighted_gale_shapley(students_obj, units_obj)
        unmatched_count = sum(1 for m in matches.values() if m is None)

        # אם מצאנו שיבוץ טוב יותר - שומרים אותו
        if unmatched_count < best_unmatched_count:
            best_unmatched_count = unmatched_count
            best_matches = matches
            best_reasons = reasons
            best_powers = {n: u.power for n, u in units_obj.items()}
            
        # אם הגענו ל-0 לא משובצים, אפשר לעצור מוקדם
        if best_unmatched_count == 0:
            break

    return (best_matches, best_reasons), best_powers