import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import os
import base64
import urllib.parse
import pandas as pd
import io
import hashlib
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# --- 1. إعدادات الصفحة الأساسية ---
st.set_page_config(page_title="منصة الأستاذ الذكي - الجزائر", page_icon="📚", layout="wide")

# تهيئة متغيرات الجلسة (Session State)
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_data" not in st.session_state:
    st.session_state.user_data = {}
if "doc_id" not in st.session_state:
    st.session_state.doc_id = None
if "generated_doc" not in st.session_state:
    st.session_state.generated_doc = False
if "doc_details" not in st.session_state:
    st.session_state.doc_details = {}

# دالة لتشفير كلمات المرور للأمان
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# --- 2. إعداد التصميم والواجهة ---
def get_base64_of_bin_file(bin_file):
    if not os.path.exists(bin_file):
        return ""
    try:
        with open(bin_file, 'rb') as f:
            data = f.read()
        return base64.b64encode(data).decode()
    except Exception:
        return ""

def set_png_as_page_bg(png_file):
    bin_str = get_base64_of_bin_file(png_file)
    if bin_str:
        bg_style = f'background-image: linear-gradient(rgba(255, 255, 255, 0.90), rgba(255, 255, 255, 0.90)), url("data:image/png;base64,{bin_str}"); background-size: cover; background-attachment: fixed;'
    else:
        bg_style = 'background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 50%, #f1f5f9 100%); background-attachment: fixed;'

    page_bg_img = f'''
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=Cairo:wght@400;600;700&display=swap');
    .stApp {{ {bg_style} }}
    .stApp, html, body, [class*="css"], label, input, select, button, p, span, th, td {{
        font-family: 'Cairo', 'Traditional Arabic', serif !important;
        direction: rtl;
    }}
    .greeting-text {{
        font-family: 'Amiri', 'Traditional Arabic', serif;
        font-size: 2.1rem; font-weight: 700; color: #166534; text-align: center; padding: 15px 0;
    }}
    .custom-box {{
        background-color: #FFFFFF; border-radius: 16px; padding: 30px;
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.05); border: 2px solid #22c55e; margin-bottom: 20px;
    }}
    .payment-card {{
        background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 10px;
        padding: 20px; margin: 15px 0; font-size: 1rem; color: #166534; line-height: 1.6;
    }}
    .vip-box {{
        background: #fffbeb; border: 2px dashed #f59e0b; border-radius: 12px;
        padding: 20px; margin: 15px 0; color: #b45309; font-weight: bold;
    }}
    .document-preview {{
        background: #ffffff; border: 2px solid #166534; border-radius: 12px;
        padding: 25px; margin-top: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        font-family: 'Traditional Arabic', serif !important; font-size: 14pt;
    }}
    </style>
    '''
    st.markdown(page_bg_img, unsafe_allow_html=True)

set_png_as_page_bg('background.jpg')

# --- 3. تهيئة الاتصال بـ Firebase Firestore (تدعم st.secrets للأمان) ---
@st.cache_resource
def init_firebase():
    try:
        if not firebase_admin._apps:
            if "firebase" in st.secrets:
                firebase_creds = dict(st.secrets["firebase"])
                cred = credentials.Certificate(firebase_creds)
                firebase_admin.initialize_app(cred)
            else:
                if os.path.exists("firebase_key.json"):
                    cred = credentials.Certificate("firebase_key.json")
                    firebase_admin.initialize_app(cred)
                else:
                    return None
        return firestore.client()
    except Exception as e:
        st.error(f"خطأ في الاتصال بقاعدة البيانات: {e}")
        return None

db = init_firebase()

ADMIN_WHATSAPP = "213672828870"
CCP_ACCOUNT = "4427739"
CCP_KEY = "34"
RIP_NUMBER = "00799999000442773972"

try:
    correct_admin_pass = st.secrets["admin_password"]
except Exception:
    correct_admin_pass = "admin123"

GENERAL_DATABASE = {
    "التعليم الابتدائي": {
        "مستويات": ["السنة الأولى ابتدائي (1AP)", "السنة الثانية ابتدائي (2AP)", "السنة الثالثة ابتدائي (3AP)", "السنة الرابعة ابتدائي (4AP)", "السنة الخامسة ابتدائي (5AP)"],
        "مواد": ["أستاذ مادة عامة", "اللغة العربية", "اللغة الفرنسية", "اللغة الإنجليزية"]
    },
    "التعليم المتوسط": {
        "مستويات": ["السنة الأولى متوسط (1AM)", "السنة الثانية متوسط (2AM)", "السنة الثالثة متوسط (3AM)", "السنة الرابعة متوسط (4AM)"],
        "مواد": ["اللغة العربية وآدابها", "الرياضيات", "العلوم الفيزيائية والتكنولوجية", "علوم الطبيعة والحياة", "التاريخ والجغرافيا", "اللغة الفرنسية", "اللغة الإنجليزية"]
    },
    "التعليم الثانوي": {
        "مستويات": ["السنة الأولى ثانوي (1AS)", "السنة الثانية ثانوي (2AS)", "السنة الثالثة ثانوي (3AS - بكالوريا)"],
        "مواد": ["الرياضيات", "علوم الطبيعة والحياة", "العلوم الفيزيائية", "الفلسفة", "اللغة العربية وآدابها", "التاريخ والجغرافيا"]
    }
}

st.markdown('<p class="greeting-text">أستاذي، أستاذتي.. التنظيم البيداغوجي الدقيق يعكس احترافية الأستاذ</p>', unsafe_allow_html=True)

# --- دالة توليد الوثيقة الرسمية بصيغة Word ---
def create_word_document(u, d):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    def add_styled_para(text, bold=False, align=WD_ALIGN_PARAGRAPH.RIGHT, size=14):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.font.name = 'Traditional Arabic'
        run.font.size = Pt(size)
        run.bold = bold
        return p

    add_styled_para("الجمهورية الجزائرية الديمقراطية الشعبية", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=16)
    add_styled_para("وزارة التربية الوطنية", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=14)
    add_styled_para(f"مديرية التربية لولاية: {u.get('district', '....................')}", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, size=14)
    add_styled_para(f"المؤسسة التعليمية: {u.get('school', '....................')}", bold=False, align=WD_ALIGN_PARAGRAPH.RIGHT, size=14)
    
    doc.add_paragraph()

    add_styled_para("بطاقة المذكرة البيداغوجية / كراس اليومية", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=16)
    doc.add_paragraph()

    table = doc.add_table(rows=5, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    table_data = [
        ("الطور التعليمي:", f"{u.get('tour', '')} - {u.get('level', '')}"),
        ("المادة والمقطع:", f"{u.get('subject', '')} | المقطع: {d['sequence']}"),
        ("الأستاذ(ة):", u.get('full_name', '')),
        ("الأسبوع التربوي:", str(d['week'])),
        ("اليوم والتاريخ والنشاط:", f"اليوم: {d['day']} | النشاط: {d['title']}")
    ]

    for i, (label, val) in enumerate(table_data):
        row = table.rows[i]
        cell_label = row.cells[1]
        cell_val = row.cells[0]
        
        cell_label.text = label
        cell_val.text = val
        
        for cell in (cell_label, cell_val):
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for run in paragraph.runs:
                    run.font.name = 'Traditional Arabic'
                    run.font.size = Pt(14)
                    run.bold = True

    doc.add_paragraph()
    
    add_styled_para(f"الكفاءة المستهدفة: {d['competency']}", bold=True, size=14)
    add_styled_para("1. الوسائل التعليمية المستعملة:", bold=True, size=14)
    add_styled_para(f"- {d['tools']}", bold=False, size=14)
    
    doc.add_paragraph()
    add_styled_para("2. سير الدرس ومراحل الإنجاز (البناء والتقويم):", bold=True, size=14)
    add_styled_para(f"- {d['notes']}", bold=False, size=14)

    doc.add_paragraph()
    doc.add_paragraph()

    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_sig = p_sig.add_run("إمضاء الأستاذ(ة):                                    إمضاء السيد المدير / الناظر:\n\n\n")
    run_sig.font.name = 'Traditional Arabic'
    run_sig.font.size = Pt(14)
    run_sig.bold = True

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- الواجهة البرمجية ---
portal_type = st.sidebar.radio("اختر البوابة:", ["بوابة الأساتذة 👨‍🏫", "لوحة تحكم المشرف 🔐"])

if portal_type == "بوابة الأساتذة 👨‍🏫":
    menu = ["تسجيل الدخول", "تسجيل حساب جديد", "استرجاع وتغيير كلمة المرور 🔑"]
    choice = st.selectbox("القائمة:", menu)

    if choice == "تسجيل حساب جديد":
        st.markdown("<div class='custom-box'>", unsafe_allow_html=True)
        st.subheader("📝 تسجيل حساب أستاذ جديد (الأسبوع الأول مجانـاً)")
        with st.form("signup_form"):
            new_user = st.text_input("اسم المستخدم:")
            new_pass = st.text_input("كلمة المرور:", type="password")
            full_name = st.text_input("الاسم واللقب الكريم:")
            phone = st.text_input("رقم الهاتف (مثال: 06xxxxxxxx):")
            district = st.text_input("المقاطعة التربوية (ولاية...):")
            school = st.text_input("اسم المؤسسة التعليمية:")
            
            tour = st.selectbox("الطور التعليمي:", list(GENERAL_DATABASE.keys()))
            level = st.selectbox("المستوى الدراسي:", GENERAL_DATABASE[tour]["مستويات"])
            subject = st.selectbox("المادة:", GENERAL_DATABASE[tour]["مواد"])
            
            submit = st.form_submit_button("حفظ البيانات والبدء الفوري 🚀")
            
            if submit:
                if new_user and new_pass and full_name and phone:
                    if db:
                        try:
                            users_ref = db.collection("teachers")
                            query = users_ref.where("username", "==", new_user).stream()
                            if list(query):
                                st.error("⚠️ اسم المستخدم مسجل مسبقاً.")
                            else:
                                teacher_data = {
                                    "username": new_user,
                                    "password": hash_password(new_pass),
                                    "full_name": full_name,
                                    "phone": phone,
                                    "district": district,
                                    "school": school,
                                    "tour": tour,
                                    "level": level,
                                    "subject": subject,
                                    "subscription_type": "تجريبي (الأسبوع الأول)",
                                    "is_vip": False
                                }
                                users_ref.add(teacher_data)
                                st.success("🎉 تم تسجيل حسابك بنجاح بأمان تام! يمكنك الآن تسجيل الدخول.")
                                st.balloons()
                        except Exception as e:
                            st.error(f"حدث خطأ أثناء التسجيل: {e}")
                    else:
                        st.error("❌ لا يوجد اتصال بقاعدة البيانات.")
                else:
                    st.warning("⚠️ يرجى ملء الحقول الإجبارية.")
        st.markdown("</div>", unsafe_allow_html=True)

    elif choice == "تسجيل الدخول":
        st.markdown("<div class='custom-box'>", unsafe_allow_html=True)
        st.subheader("🔐 تسجيل الدخول للحساب")
        log_user = st.text_input("اسم المستخدم:")
        log_pass = st.text_input("كلمة المرور:", type="password")
        
        if st.button("دخول للحساب 🔍"):
            if db:
                try:
                    users_ref = db.collection("teachers")
                    hashed_input_pass = hash_password(log_pass)
                    query = users_ref.where("username", "==", log_user).where("password", "==", hashed_input_pass).stream()
                    user_found = None
                    doc_id = None
                    for doc in query:
                        user_found = doc.to_dict()
                        doc_id = doc.id
                        
                    if user_found:
                        st.session_state.logged_in = True
                        st.session_state.user_data = user_found
                        st.session_state.doc_id = doc_id
                        st.success(f"مرحباً بك أستاذ(ة) {user_found['full_name']} 🎓")
                        st.rerun()
                    else:
                        st.error("❌ بيانات الدخول غير صحيحة.")
                except Exception as e:
                    st.error(f"خطأ في الاتصال: {e}")
            else:
                st.error("❌ لا يوجد اتصال بقاعدة البيانات.")
        st.markdown("</div>", unsafe_allow_html=True)

    elif choice == "استرجاع وتغيير كلمة المرور 🔑":
        st.markdown("<div class='custom-box'>", unsafe_allow_html=True)
        st.subheader("🔑 استرجاع وتحديث كلمة المرور الذاتي")
        st.write("أدخل اسم المستخدم ورقم الهاتف المسجل في حسابك للتحقق من هويتك وتعيين كلمة مرور جديدة فوراً:")
        
        with st.form("reset_pass_form"):
            r_user = st.text_input("اسم المستخدم:")
            r_phone = st.text_input("رقم الهاتف المسجل في الحساب:")
            new_pass1 = st.text_input("كلمة المرور الجديدة:", type="password")
            new_pass2 = st.text_input("تأكيد كلمة المرور الجديدة:", type="password")
            
            reset_submit = st.form_submit_button("تحديث كلمة المرور 🔄")
            
            if reset_submit:
                if r_user and r_phone and new_pass1 and new_pass2:
                    if new_pass1 != new_pass2:
                        st.error("❌ كلمتا المرور غير متطابقتين.")
                    else:
                        if db:
                            try:
                                users_ref = db.collection("teachers")
                                query = users_ref.where("username", "==", r_user).where("phone", "==", r_phone).stream()
                                doc_to_update = None
                                for doc in query:
                                    doc_to_update = doc.id
                                    
                                if doc_to_update:
                                    users_ref.document(doc_to_update).update({
                                        "password": hash_password(new_pass1)
                                    })
                                    st.success("🎉 تم تحديث كلمة المرور بنجاح! يمكنك الآن تسجيل الدخول بكلمة المرور الجديدة.")
                                else:
                                    st.error("❌ البيانات المدخلة (اسم المستخدم أو رقم الهاتف) غير مطابقة لأي حساب مسجل.")
                            except Exception as e:
                                st.error(f"حدث خطأ: {e}")
                        else:
                            st.error("❌ لا يوجد اتصال بقاعدة البيانات.")
                else:
                    st.warning("⚠️ يرجى ملء كافة الحقول المطلوبة.")
                    
        st.markdown("---")
        st.write("أو يمكنك التواصل مباشرة مع المشرف للمساعدة عبر الواتساب:")
        wa_recovery_link = f"https://wa.me/{ADMIN_WHATSAPP}?text=" + urllib.parse.quote("السلام عليكم المشرف، أرغب في مساعدة لاستعادة حسابي بمنصة الأستاذ الذكي.")
        st.markdown(f'<a href="{wa_recovery_link}" target="_blank" style="display:block; background:#25D366; color:white; padding:10px; text-align:center; border-radius:10px; text-decoration:none; font-weight:bold; margin-top: 5px;">💬 تواصل مع المشرف عبر الواتساب</a>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # لوحة تحكم الأستاذ بعد تسجيل الدخول
    if st.session_state.get("logged_in", False):
        u = st.session_state.user_data
        is_vip = u.get("is_vip", False)
        sub_type = u.get("subscription_type", "تجريبي (الأسبوع الأول)")
        
        st.markdown(f"### لوحة تحكم الأستاذ: {u.get('full_name', '')} | الباقة الحالية: {sub_type}")
        
        with st.form("journal_form"):
            st.write("إعداد الوثيقة والبطاقة البيداغوجية الرسمية:")
            col1, col2 = st.columns(2)
            with col1:
                sel_day = st.selectbox("اليوم:", ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس"])
                week_num = st.number_input("رقم الأسبوع التربوي:", 1, 35, 1)
            with col2:
                sequence_name = st.text_input("مقطع التعلم / الميدان:", "المقطع الأول: ...")
                lesson_title = st.text_input("عنوان الدرس / النشاط:")
            
            competency = st.text_input("الكفاءة المستهدفة (الختامية / المستعرضة):", "التحكم في المفاهيم الأساسية وتوظيفها...")
            tools = st.text_input("الوسائل التعليمية المعتمدة:", "الكتاب المدرسي، السبورة، الوسائل الإيضاحية")
            notes = st.text_area("الملاحظات والتوجيهات البيداغوجية وسير الدرس:", "التقويم التشخيصي، وضعية الانطلاق، بناء المفهوم، الاستثمار والتقويم التحصيلي.")
            
            gen_submit = st.form_submit_button("معاينة وتوليد الوثيقة الرسمية 📄")

            if gen_submit:
                st.session_state.generated_doc = True
                st.session_state.doc_details = {
                    "day": sel_day,
                    "week": week_num,
                    "sequence": sequence_name,
                    "title": lesson_title,
                    "competency": competency,
                    "tools": tools,
                    "notes": notes
                }

        is_free_week = (st.session_state.doc_details.get("week", 1) == 1) if st.session_state.generated_doc else True

        if st.session_state.generated_doc:
            if is_free_week or is_vip:
                d = st.session_state.doc_details
                word_file = create_word_document(u, d)
                
                st.markdown(f"""
                <div class="document-preview">
                    <h3 style="color: #166534; text-align: center;">الجمهورية الجزائرية الديمقراطية الشعبية - وزارة التربية الوطنية</h3>
                    <hr>
                    <p><b>المؤسسة:</b> {u.get('school', 'غير محدد')} &nbsp;&nbsp;&nbsp;&nbsp; <b>المقاطعة:</b> {u.get('district', 'غير محدد')}</p>
                    <p><b>الأستاذ(ة):</b> {u.get('full_name', '')} &nbsp;&nbsp;&nbsp;&nbsp; <b>الطور:</b> {u.get('tour', '')} &nbsp;&nbsp;&nbsp;&nbsp; <b>المستوى:</b> {u.get('level', '')}</p>
                    <p><b>المادة:</b> {u.get('subject', '')} &nbsp;&nbsp;&nbsp;&nbsp; <b>الأسبوع:</b> {d['week']} &nbsp;&nbsp;&nbsp;&nbsp; <b>اليوم:</b> {d['day']}</p>
                    <p><b>المقطع:</b> {d['sequence']} &nbsp;&nbsp;&nbsp;&nbsp; <b>النشاط:</b> {d['title']}</p>
                    <hr>
                    <p><b>الكفاءة المستهدفة:</b> {d['competency']}</p>
                    <p><b>الوسائل المستعملة:</b> {d['tools']}</p>
                    <p><b>الملاحظات البيداغوجية:</b> {d['notes']}</p>
                    <br>
                    <p style="text-align: left;"><b>إمضاء الأستاذ(ة):</b> ........................</p>
                </div>
                """, unsafe_allow_html=True)
                
                st.success("✅ تمت مطابقة الوثيقة بنجاح وفق المعايير الرسمية وجاهزة للتحميل بصيغة Word!")
                
                st.download_button(
                    label="📥 تحميل الوثيقة الرسمية كملف Word (Traditional Arabic 14)",
                    data=word_file,
                    file_name=f"مذكرة_رسمية_الأسبوع_{d['week']}_{u.get('full_name', 'أستاذ')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                if sub_type == "اشتراك سنوي (VIP)":
                    st.markdown("""
                    <div class="vip-box">
                        🎁 <b>هدايا الاشتراك السنوي الخاصة بك:</b><br>
                        لديك الصلاحية الكاملة لتحميل <b>حقيبة الملفات البيداغوجية والإدارية المنظمة ومرتبة لكامل السنة</b> (تتضمن الكشوف، الحجرات، قوائم التلاميذ، ومختلف السجلات الإدارية المطلوبة من طرف لجان التفتيش بوزارة التربية الوطنية).
                    </div>
                    """, unsafe_allow_html=True)
                
            else:
                st.warning(f"🔒 الأسبوع رقم {st.session_state.doc_details.get('week')} يتطلب اشتراكاً مدفوعاً.")
                
                st.markdown(f"""
                <div class="payment-card">
                    <b>💳 اختر باقة الاشتراك المناسبة لتفعيل حسابك عبر بريد الجزائر (CCP):</b><br><br>
                    🔹 <b>الاشتراك الشهري:</b> 1,000 دج<br>
                    🔹 <b>الاشتراك الفصلي:</b> 2,500 دج<br>
                    🌟 <b>الاشتراك السنوي:</b> 6,000 دج <i>(يتضمن منح حقيبة ملفات منظمة ومرتبة لكامل السنة ولكل الوثائق البيداغوجية والإدارية حسب وزارة التربية الوطنية)</i><br><br>
                    - <b>رقم الحساب (CCP):</b> <code>{CCP_ACCOUNT}</code> (المفتاح: <b>{CCP_KEY}</b>)<br>
                    - <b>رقم التعريف البريدي (RIP):</b> <code>{RIP_NUMBER}</code><br><br>
                    <i>بعد إتمام الدفع، أرسل وصل الدفع عبر الواتساب مع تحديد نوع الاشتراك المطلوب لتفعيل حسابك فوراً.</i>
                </div>
                """, unsafe_allow_html=True)

                wa_msg = f"السلام عليكم المشرف، قمت بتسديد اشتراك بمنصة الأستاذ الذكي وأرغب في تفعيل حسابي.\n\n👤 اسم الأستاذ(ة): {u.get('full_name', '')}\n📱 الهاتف: {u.get('phone', '')}\n🏫 المؤسسة: {u.get('school', 'غير محدد')}"
                wa_link = f"https://wa.me/{ADMIN_WHATSAPP}?text=" + urllib.parse.quote(wa_msg)
                
                st.markdown(f'<a href="{wa_link}" target="_blank" style="display:block; background:#25D366; color:white; padding:14px; text-align:center; border-radius:10px; text-decoration:none; font-weight:bold; font-size: 1.1rem; margin-top: 15px;">💬 تواصل عبر الواتساب لتأكيد الباقة والتفعيل الفوري</a>', unsafe_allow_html=True)

        if st.button("تسجيل الخروج 🚪"):
            st.session_state.logged_in = False
            st.session_state.user_data = {}
            st.session_state.doc_id = None
            st.session_state.generated_doc = False
            st.session_state.doc_details = {}
            st.rerun()

elif portal_type == "لوحة تحكم المشرف 🔐":
    st.markdown("<div class='custom-box'>", unsafe_allow_html=True)
    st.subheader("🛠️ بوابة تسجيل دخول المشرف العام")
    admin_pass_input = st.text_input("كلمة مرور المشرف السرية:", type="password")
    
    if admin_pass_input == correct_admin_pass:
        st.success("تم الدخول بنجاح إلى لوحة الإدارة ✅")
        
        if db:
            try:
                teachers_ref = db.collection("teachers").stream()
                teachers_list = []
                vip_count = 0
                for doc in teachers_ref:
                    t_data = doc.to_dict()
                    t_data["id"] = doc.id
                    teachers_list.append(t_data)
                    if t_data.get("is_vip", False):
                        vip_count += 1
                
                st.markdown("### 📊 إحصائيات عامة للمنصة:")
                col_a, col_b, col_c = st.columns(3)
                col_a.metric("إجمالي الأساتذة المسجلين", len(teachers_list))
                col_b.metric("المشتركون النشطون (VIP)", vip_count)
                col_c.metric("المستخدمون في الفترة التجريبية", len(teachers_list) - vip_count)
                
                st.markdown("---")
                st.markdown("### 📋 جدول جميع الأساتذة المسجلين:")
                
                if not teachers_list:
                    st.info("لا يوجد أساتذة مسجلين بعد.")
                else:
                    table_data = []
                    for t in teachers_list:
                        table_data.append({
                            "الاسم الكامل": t.get("full_name"),
                            "رقم الهاتف": t.get("phone"),
                            "اسم المستخدم": t.get("username"),
                            "المادة والمستوى": f"{t.get('subject')} - {t.get('level')}",
                            "المؤسسة": t.get("school", ""),
                            "نوع الباقة": t.get("subscription_type", "تجريبي")
                        })
                    
                    df = pd.DataFrame(table_data)
                    st.dataframe(df, use_container_width=True)
                    
                    st.markdown("---")
                    st.markdown("### ⚙️ إدارة وتفعيل باقات الأساتذة الفردية:")
                    
                    for t in teachers_list:
                        with st.expander(f"التحكم في الأستاذ: {t.get('full_name', '')} (الهاتف: {t.get('phone', '')})"):
                            st.write(f"👤 **اسم المستخدم:** `{t.get('username', '')}` | 📍 **المقاطعة:** {t.get('district', '')}")
                            
                            current_vip = t.get("is_vip", False)
                            new_vip_status = st.checkbox("تفعيل الحساب (VIP)", value=current_vip, key=f"vip_{t['id']}")
                            
                            new_sub_type = st.selectbox(
                                "اختر باقة الاشتراك المعتمدة:",
                                ["تجريبي (الأسبوع الأول)", "اشتراك شهري (1000 دج)", "اشتراك فصلي (2500 دج)", "اشتراك سنوي (VIP)"],
                                index=0 if not current_vip else 3,
                                key=f"sub_{t['id']}"
                            )
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.button(f"حفظ التعديل والباقة", key=f"save_{t['id']}"):
                                    db.collection("teachers").document(t['id']).update({
                                        "is_vip": new_vip_status,
                                        "subscription_type": new_sub_type
                                    })
                                    st.success("تم تحديث باقة الأستاذ بنجاح!")
                                    st.rerun()
                            with col2:
                                if st.button(f"🗑️ حذف الحساب", key=f"del_{t['id']}", type="primary"):
                                    db.collection("teachers").document(t['id']).delete()
                                    st.warning("تم حذف الحساب نهائياً.")
                                    st.rerun()
            except Exception as e:
                st.error(f"خطأ أثناء جلب بيانات الأساتذة من قاعدة البيانات: {e}")
        else:
            st.error("❌ لا يوجد اتصال بقاعدة البيانات.")
    elif admin_pass_input != "":
        st.error("كلمة مرور المشرف خاطئة.")
    st.markdown("</div>", unsafe_allow_html=True)
