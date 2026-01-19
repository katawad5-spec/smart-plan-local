import streamlit as st
import pandas as pd
import altair as alt
import streamlit.components.v1 as components
import io
import json
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==========================================
# ⚙️ 1. System Config
# ==========================================
st.set_page_config(page_title="ระบบจัดทำแผนฯ (Final Master)", layout="wide", initial_sidebar_state="expanded")

# --- CSS: จัดการหน้าจอและงานพิมพ์ (หัวใจสำคัญ) ---
st.markdown("""
<style>
    /* Font */
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap');
    
    /* Dashboard Cards */
    .metric-card { background-color: #f8f9fa; border: 1px solid #dee2e6; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    
    /* กระดาษ A4 (Preview) */
    .a4-paper { 
        background-color: white; padding: 2.54cm; margin: 0 auto 20px auto; 
        width: 210mm; min-height: 297mm; box-shadow: 0 4px 8px rgba(0,0,0,0.2); 
        color: black; font-family: 'Sarabun', sans-serif; font-size: 16pt; line-height: 1.5;
    }
    
    /* ซ่อนเมนูเวลาสั่งปริ้นท์ */
    @media print {
        [data-testid="stSidebar"], [data-testid="stHeader"], .stApp > header, .stApp > footer, .no-print { display: none !important; }
        .stApp { background: white; margin: 0; padding: 0; }
        .block-container { padding: 0 !important; max-width: 100% !important; }
        .a4-paper { box-shadow: none; margin: 0; width: 100%; page-break-after: always; }
    }
    
    /* ตารางใน A4 (ตามระเบียบงานสารบรรณ) */
    table { width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 10px; }
    th, td { border: 1px solid black; padding: 5px; text-align: left; vertical-align: top; font-size: 14pt; }
    th { text-align: center; background-color: #f0f0f0; font-weight: bold; }
    
    /* จัดหน้า */
    h1 { font-size: 24pt; font-weight: bold; text-align: center; margin-bottom: 20px; }
    h2 { font-size: 20pt; font-weight: bold; margin-top: 20px; }
    .indent { text-indent: 1cm; text-align: justify; }
</style>
""", unsafe_allow_html=True)

# --- Master Data (ข้อมูลหลัก) ---
NATIONAL_STRAT_LIST = [
    "1. ด้านความมั่นคง", "2. ด้านการสร้างความสามารถในการแข่งขัน", 
    "3. ด้านการพัฒนาและเสริมสร้างศักยภาพทรัพยากรมนุษย์", "4. ด้านการสร้างโอกาสและความเสมอภาคทางสังคม", 
    "5. ด้านการสร้างการเติบโตบนคุณภาพชีวิตที่เป็นมิตรต่อสิ่งแวดล้อม", "6. ด้านการปรับสมดุลและพัฒนาระบบการบริหารจัดการภาครัฐ"
]
PROVINCIAL_STRAT_LIST = [
    "1. การค้า/การลงทุน/ท่องเที่ยว", "2. เกษตรอัจฉริยะ", "3. คุณภาพชีวิต/สังคมน่าอยู่", 
    "4. ทรัพยากรธรรมชาติ", "5. ความมั่นคง"
]
STRAT_LIST = [
    "1. ด้านโครงสร้างพื้นฐาน", "2. ด้านเศรษฐกิจและท่องเที่ยว", 
    "3. ด้านคุณภาพชีวิตและสังคม", "4. ด้านทรัพยากรธรรมชาติและสิ่งแวดล้อม", 
    "5. ด้านการบริหารจัดการบ้านเมืองที่ดี"
]
ORG_DIVISIONS = [
    "สำนักปลัด (อบต.หนองแสง)", "กองคลัง (อบต.หนองแสง)", "กองช่าง (อบต.หนองแสง)", 
    "อบจ.อุดรธานี", "กรมทางหลวง", "กรมทางหลวงชนบท", "การไฟฟ้าส่วนภูมิภาค", "อื่นๆ"
]
TOPICS_P1 = {1: "กายภาพ", 2: "ขอบเขต", 3: "ประชากร", 4: "คมนาคมทางบก", 5: "โลจิสติกส์", 6: "คมนาคมทางน้ำ", 7: "คมนาคมทางอากาศ", 8: "ขนส่งสาธารณะ", 9: "เมืองอัจฉริยะ", 10: "ดิจิทัล", 11: "การศึกษา", 12: "อัตลักษณ์", 13: "ศาสนา/วัฒนธรรม", 14: "ภูมิปัญญา", 15: "สาธารณสุข", 16: "สังคมสงเคราะห์", 17: "ความปลอดภัย", 18: "ยาเสพติด", 19: "สาธารณภัย", 20: "ประชาสังคม", 21: "ไฟฟ้า", 22: "บำบัดน้ำเสีย", 23: "ขยะ", 24: "ตลาด", 25: "แหล่งน้ำ", 26: "ทรัพยากรธรรมชาติ", 27: "ป่าชุมชน", 28: "อาชีพ", 29: "เกษตร", 30: "ประมง/ปศุสัตว์", 31: "ท่องเที่ยว", 32: "อุตสาหกรรม", 33: "พาณิชย์", 34: "แรงงาน", 35: "กีฬา", 36: "รายได้", 37: "อื่น ๆ"}

# --- Initial State ---
if 'projects' not in st.session_state: st.session_state.projects = []
# Mapping อัตโนมัติ (Default)
if 'strat_mapping' not in st.session_state:
    st.session_state.strat_mapping = {s: {"nat": NATIONAL_STRAT_LIST[0], "prov": PROVINCIAL_STRAT_LIST[0]} for s in STRAT_LIST}
if 'general_info' not in st.session_state:
    st.session_state.general_info = {f"p1_{i}": "-" for i in range(1, 38)}
    st.session_state.general_info.update({"local_name": "องค์การบริหารส่วนตำบลหนองแสง", "vision": "-", "policy": "-", "linkage": "-", "strat_issues": "-", "part4": "-"})

# ==========================================
# 🧠 Logic Functions (สมองของระบบ)
# ==========================================
def clean_text(text): 
    if not isinstance(text, str): return text
    return re.sub(' +', ' ', text.strip()).replace(" ,", ",").replace(" .", ".")

def to_thai_num(n): 
    return str(n).translate(str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙"))

def check_duplicate(name):
    return any(p['name'].strip() == name.strip() for p in st.session_state.projects)

def smart_input(label, key_base, suggestions):
    sel, txt = f"s_{key_base}", f"t_{key_base}"
    if txt not in st.session_state: st.session_state[txt] = ""
    def chg(): 
        if st.session_state[sel] != "- เลือกตัวอย่าง -": st.session_state[txt] = st.session_state[sel]
    c1, c2 = st.columns([1,3])
    with c1: st.selectbox("💡", suggestions, key=sel, on_change=chg, label_visibility="collapsed")
    with c2: return st.text_input(label, key=txt)

# ==========================================
# 📥 Excel Logic (Robust Import)
# ==========================================
def create_excel_template():
    df = pd.DataFrame(columns=["ประเภท", "เลขประเด็น(1-5)", "แผนงาน", "ชื่อโครงการ", "วัตถุประสงค์", "เป้าหมาย", "งบ71", "งบ72", "งบ73", "งบ74", "งบ75", "ตัวชี้วัด", "ผลลัพธ์", "หน่วยงาน"])
    df.loc[0] = ["ปกติ", 1, "เคหะฯ", "ก่อสร้างถนน...", "สัญจร", "500 ม.", 500000, 0, 0, 0, 0, "1 สาย", "สะดวก", "กองช่าง (อบต.หนองแสง)"]
    output = io.BytesIO(); with pd.ExcelWriter(output, engine='xlsxwriter') as writer: df.to_excel(writer, index=False); return output.getvalue()

def process_excel(file):
    try:
        df = pd.read_excel(file); df.columns = df.columns.str.strip(); df = df.fillna(0)
        new_data = []; skipped = 0
        names = set(p['name'].strip() for p in st.session_state.projects)
        for _, row in df.iterrows():
            name = clean_text(str(row.get("ชื่อโครงการ", "")))
            if not name or "ตัวอย่าง" in name: continue
            if name in names: skipped += 1; continue
            try: s_idx = int(row.get("เลขประเด็น(1-5)", 1))
            except: s_idx = 1
            strat_val = STRAT_LIST[s_idx-1] if 0<s_idx<=5 else STRAT_LIST[0]
            
            # Map หน่วยงานให้ตรงกับที่มีในระบบ ถ้าไม่ตรงให้คงเดิมไว้
            raw_owner = clean_text(str(row.get("หน่วยงาน","")))
            owner = raw_owner if raw_owner in ORG_DIVISIONS else raw_owner

            new_data.append({
                "type": str(row.iloc[0]), "strat": strat_val,
                "name": name, "obj": clean_text(str(row.get("วัตถุประสงค์",""))), "target": clean_text(str(row.get("เป้าหมาย",""))),
                "b1": float(row.get("งบ71",0)), "b2": float(row.get("งบ72",0)), "b3": float(row.get("งบ73",0)),
                "b4": float(row.get("งบ74",0)), "b5": float(row.get("งบ75",0)),
                "kpi": clean_text(str(row.get("ตัวชี้วัด",""))), "result": clean_text(str(row.get("ผลลัพธ์",""))), "owner": owner
            })
            names.add(name)
        st.session_state.projects.extend(new_data); return len(new_data), skipped
    except: return 0, 0

# ==========================================
# 📄 Custom Print Generator (HTML Output)
# ==========================================
def generate_print_html(options):
    data = st.session_state.general_info
    projects = st.session_state.projects
    df = pd.DataFrame(projects) if projects else pd.DataFrame()
    
    html = ""
    # 1. หน้าปก
    if options.get('cover'):
        html += f"""
        <div class='a4-paper'>
            <br><br><br><br><br>
            <h1>แผนพัฒนาท้องถิ่น (พ.ศ. ๒๕๗๑ - ๒๕๗๕)</h1>
            <br><br>
            <h1>{data['local_name']}</h1>
            <div style="text-align:center; margin-top:100px; font-size:16pt;">
                งานวิเคราะห์นโยบายและแผน<br>สำนักปลัด {data['local_name']}
            </div>
        </div>
        """
    
    # 2. ส่วนที่ 1
    if options.get('p1'):
        html += "<div class='a4-paper'><h2>ส่วนที่ ๑ สภาพทั่วไปและข้อมูลพื้นฐาน</h2>"
        for i in range(1, 38):
            val = data.get(f"p1_{i}", "-")
            if val != "-" and val != "": 
                html += f"<p><b>๑.{to_thai_num(i)} ด้าน{TOPICS_P1[i]}</b></p><p class='indent'>{val}</p>"
        html += "</div>"
        
    # 3. ส่วนที่ 2 (เชื่อมโยง)
    if options.get('p2'):
        html += f"""
        <div class='a4-paper'>
            <h2>ส่วนที่ ๒ ประเด็นการพัฒนาท้องถิ่น</h2>
            <p><b>๒.๑ วิสัยทัศน์:</b> {data['vision']}</p>
            <p><b>๒.๒ พันธกิจ:</b> {data['policy']}</p>
            <p><b>๒.๓ ความเชื่อมโยง:</b> {data['linkage']}</p>
            <p><b>๒.๔ ยุทธศาสตร์การพัฒนา:</b> {data['strat_issues']}</p>
        </div>
        """
        
    # 4. ส่วนที่ 3 (โครงการ)
    if options.get('p3') and not df.empty:
        # ผ.01 (สรุป)
        html += """<div class='a4-paper'><h2>ส่วนที่ ๓ การนำแผนไปสู่การปฏิบัติ</h2><p><b>๓.๑ บัญชีสรุปโครงการ (ผ.๐๑)</b></p>"""
        html += "<table><thead><tr><th>ประเด็นการพัฒนา</th><th>๒๕๗๑</th><th>๒๕๗๒</th><th>๒๕๗๓</th><th>๒๕๗๔</th><th>๒๕๗๕</th><th>รวม</th></tr></thead><tbody>"
        grp = df.groupby('strat')[['b1','b2','b3','b4','b5']].sum().reset_index()
        for _, r in grp.iterrows():
            total = r['b1']+r['b2']+r['b3']+r['b4']+r['b5']
            html += f"<tr><td>{r['strat']}</td><td align='right'>{r['b1']:,.0f}</td><td align='right'>{r['b2']:,.0f}</td><td align='right'>{r['b3']:,.0f}</td><td align='right'>{r['b4']:,.0f}</td><td align='right'>{r['b5']:,.0f}</td><td align='right'>{total:,.0f}</td></tr>"
        html += "</tbody></table></div>"
        
        # ผ.02 (รายละเอียด)
        html += "<div class='a4-paper'><h2>แบบ ผ.๐๒ บัญชีรายละเอียดโครงการ</h2>"
        for strat in sorted(df['strat'].unique()):
            html += f"<h4>{strat}</h4>"
            html += "<table><thead><tr><th width='5%'>ที่</th><th width='25%'>โครงการ</th><th width='20%'>เป้าหมาย</th><th width='15%'>งบประมาณ</th><th width='15%'>ตัวชี้วัด</th><th width='10%'>หน่วยงาน</th></tr></thead><tbody>"
            sub = df[df['strat'] == strat]
            for idx, row in enumerate(sub.to_dict('records')):
                total = row['b1']+row['b2']+row['b3']+row['b4']+row['b5']
                html += f"<tr><td align='center'>{to_thai_num(idx+1)}</td><td>{row['name']}</td><td>{row['target']}</td><td align='right'>{total:,.0f}</td><td>{row['kpi']}</td><td>{row['owner']}</td></tr>"
            html += "</tbody></table>"
        html += "</div>"
        
    return html

# ==========================================
# 🖥️ UI Application (เมนูและการแสดงผล)
# ==========================================
with st.sidebar:
    st.title("🗂️ เมนูหลัก")
    page = st.radio("เลือกขั้นตอน:", 
        ["1. ข้อมูลทั่วไป (ส่วนที่ 1)", 
         "2. กำหนดความเชื่อมโยง (ส่วนที่ 2)", 
         "3. บัญชีโครงการ (ส่วนที่ 3)", 
         "4. สรุปผล (Dashboard)", 
         "5. พิมพ์รายงาน (Print)"])
    st.markdown("---")
    if st.button("🗑️ ล้างข้อมูลใหม่ (Reset)"): 
        st.session_state.projects=[]; st.rerun()

# --- Page 1: General Info ---
if page == "1. ข้อมูลทั่วไป (ส่วนที่ 1)":
    st.title("📝 ส่วนที่ 1: สภาพทั่วไปและข้อมูลพื้นฐาน")
    with st.form("p1"):
        st.session_state.general_info['local_name'] = st.text_input("ชื่อ อปท.", st.session_state.general_info.get('local_name',''))
        t1, t2 = st.tabs(["กายภาพ/สังคม (1-20)", "เศรษฐกิจ/อื่น (21-37)"])
        with t1:
            for i in range(1, 11): k=f"p1_{i}"; st.session_state.general_info[k] = st.text_area(f"ด้านที่ {i} {TOPICS_P1[i]}", st.session_state.general_info.get(k,""), height=70)
        with t2:
            st.write("(กรอกด้านที่ 21-37 ที่นี่...)")
        st.form_submit_button("บันทึก")

# --- Page 2: Strategy Mapping (Logic Core) ---
elif page == "2. กำหนดความเชื่อมโยง (ส่วนที่ 2)":
    st.title("🔗 ส่วนที่ 2: เชื่อมโยงยุทธศาสตร์")
    st.warning("⚠️ **สำคัญ:** โปรดจับคู่ 'ยุทธศาสตร์ อปท.' กับ 'ยุทธศาสตร์ชาติ/จังหวัด' ให้ครบ เพื่อให้ Dashboard แสดงผลถูกต้อง")
    
    with st.form("mapping"):
        for local in STRAT_LIST:
            st.markdown(f"**{local}**")
            c1, c2 = st.columns(2)
            cur_nat = st.session_state.strat_mapping[local]['nat']
            cur_prov = st.session_state.strat_mapping[local]['prov']
            
            with c1: new_nat = st.selectbox(f"ยุทธศาสตร์ชาติ", NATIONAL_STRAT_LIST, index=NATIONAL_STRAT_LIST.index(cur_nat), key=f"n_{local}")
            with c2: new_prov = st.selectbox(f"แผนจังหวัด", PROVINCIAL_STRAT_LIST, index=PROVINCIAL_STRAT_LIST.index(cur_prov), key=f"p_{local}")
            
            st.session_state.strat_mapping[local] = {"nat": new_nat, "prov": new_prov}
            st.markdown("---")
        st.form_submit_button("💾 บันทึกความเชื่อมโยง")

# --- Page 3: Projects (Operation) ---
elif page == "3. บัญชีโครงการ (ส่วนที่ 3)":
    st.title("🏗️ ส่วนที่ 3: บัญชีโครงการ (ผ.02)")
    
    with st.expander("➕ เพิ่มโครงการ (Manual)", expanded=True):
        c1, c2 = st.columns([1,1])
        with c1:
            with st.form("add"):
                pt = st.selectbox("ประเภท", ["ปกติ", "เกินศักยภาพ", "อุดหนุน"])
                st_iss = st.selectbox("ยุทธศาสตร์ อปท. (ระบบเชื่อมโยงให้เอง)", STRAT_LIST)
                nm = st.text_input("ชื่อโครงการ (ห้ามซ้ำ)")
                obj = smart_input("วัตถุประสงค์", "obj", ["เพื่อการสัญจรสะดวก", "เพื่อป้องกันน้ำท่วม"])
                tgt = smart_input("เป้าหมาย", "tgt", ["กว้าง 5.00 เมตร", "จำนวน 1 แห่ง"])
                b1 = st.number_input("งบประมาณ 2571", step=10000)
                own = st.selectbox("หน่วยงานรับผิดชอบ", ORG_DIVISIONS)
                
                if st.form_submit_button("บันทึก"):
                    if check_duplicate(nm): 
                        st.error("❌ ชื่อโครงการซ้ำ! มีในระบบแล้ว")
                    elif not nm:
                        st.error("❌ กรุณาใส่ชื่อโครงการ")
                    else:
                        st.session_state.projects.append({"type":pt, "strat":st_iss, "name":clean_text(nm), "obj":clean_text(obj), "target":clean_text(tgt), "b1":b1, "b2":0, "b3":0, "b4":0, "b5":0, "kpi":"-", "result":"-", "owner":own})
                        st.success("✅ บันทึกแล้ว")
                        st.rerun()
        with c2:
            st.info("📥 นำเข้าจาก Excel")
            upl = st.file_uploader("Upload Excel", type=['xlsx'])
            if upl and st.button("Import"):
                add, skip = process_excel(upl)
                if add > 0: st.success(f"เพิ่ม {add} รายการ")
                if skip > 0: st.warning(f"ข้ามที่ซ้ำ {skip} รายการ")
                st.rerun()
            st.download_button("โหลด Template", create_excel_template(), "Form_Standard.xlsx")

    if st.session_state.projects:
        st.markdown("---")
        st.subheader(f"📋 รายการโครงการ ({len(st.session_state.projects)})")
        
        # Data Editor (แก้ไข/ลบ ได้ในตารางเลย)
        df = pd.DataFrame(st.session_state.projects)
        edited_df = st.data_editor(
            df, 
            use_container_width=True, 
            num_rows="dynamic", 
            key="editor",
            column_config={
                "b1": st.column_config.NumberColumn("งบ 71", format="%d"),
                "owner": st.column_config.SelectboxColumn("หน่วยงาน", options=ORG_DIVISIONS)
            }
        )
        if not df.equals(edited_df):
            st.session_state.projects = edited_df.to_dict('records'); st.rerun()

# --- Page 4: Dashboard (Analysis) ---
elif page == "4. สรุปผล (Dashboard)":
    st.title("📊 สรุปผล & ความเชื่อมโยง")
    if not st.session_state.projects: st.info("ไม่มีข้อมูลโครงการ"); st.stop()
    
    # 1. KPI Cards
    df = pd.DataFrame(st.session_state.projects)
    total_budget = df[['b1','b2','b3','b4','b5']].sum(axis=1).sum()
    c1, c2, c3 = st.columns(3)
    c1.metric("จำนวนโครงการ", f"{len(df)} โครงการ")
    c2.metric("งบประมาณรวม", f"{total_budget:,.0f} บาท")
    c3.metric("หน่วยงาน", f"{len(df['owner'].unique())} แห่ง")
    
    st.markdown("---")
    
    # 2. Logic Mapping (ดึงค่าจากหน้า 2 มาคำนวณ)
    mapping = st.session_state.strat_mapping
    df['nat_strat'] = df['strat'].apply(lambda x: mapping[x]['nat'])
    
    # 3. Charts
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("🇹🇭 สัดส่วนยุทธศาสตร์ชาติ (Linkage)")
        nat_count = df['nat_strat'].value_counts().reset_index()
        nat_count.columns = ['ด้าน', 'จำนวน']
        
        chart = alt.Chart(nat_count).mark_bar().encode(
            x=alt.X('จำนวน', title='จำนวนโครงการ'),
            y=alt.Y('ด้าน', sort='-x', title=None),
            color=alt.value('#1f77b4'),
            tooltip=['ด้าน','จำนวน']
        ).properties(height=300)
        st.altair_chart(chart, use_container_width=True)
        
    with col_b:
        st.subheader("🏗️ สัดส่วนหน่วยงาน")
        own_count = df['owner'].value_counts().reset_index()
        own_count.columns = ['หน่วยงาน', 'จำนวน']
        chart = alt.Chart(own_count).mark_arc().encode(
            theta=alt.Theta("จำนวน"),
            color=alt.Color("หน่วยงาน"),
            tooltip=["หน่วยงาน", "จำนวน"]
        )
        st.altair_chart(chart, use_container_width=True)

# --- Page 5: Print (Custom Output) ---
elif page == "5. พิมพ์รายงาน (Print)":
    st.title("🖨️ เลือกพิมพ์ตามสั่ง (Custom Print)")
    
    c1, c2 = st.columns([1, 3])
    with c1:
        st.write("<b>เลือกส่วนที่จะพิมพ์:</b>", unsafe_allow_html=True)
        opt_cover = st.checkbox("ปกหน้า", True)
        opt_p1 = st.checkbox("ส่วนที่ 1 (สภาพทั่วไป)", True)
        opt_p2 = st.checkbox("ส่วนที่ 2 (เชื่อมโยง)", True)
        opt_p3 = st.checkbox("ส่วนที่ 3 (โครงการ)", True)
        
        # ปุ่ม Javascript Print
        components.html("""<button onclick="window.print()" style="background:#28a745;color:white;padding:15px;width:100%;border:none;border-radius:5px;font-size:18px;cursor:pointer;font-weight:bold;">🖨️ สั่งพิมพ์ทันที</button>""", height=70)
        st.info("💡 **ทริค:** กดปุ่มเขียว แล้วเครื่องจะสั่งปริ้นท์เฉพาะส่วนที่คุณเลือก (หน้าจอด้านขวา)")

    with c2:
        # Generate Preview HTML
        html = generate_print_html({'cover': opt_cover, 'p1': opt_p1, 'p2': opt_p2, 'p3': opt_p3})
        if html: st.markdown(html, unsafe_allow_html=True)
        else: st.warning("กรุณาเลือกหัวข้อทางซ้าย")
