import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    # 16:9 Widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Theme Palette
    C_PRIMARY = RGBColor(233, 59, 103)      # #E93B67 (Warm Rose / Swasth Brand)
    C_SECONDARY = RGBColor(74, 111, 165)    # #4A6FA5 (Deep Blue)
    C_DARK = RGBColor(15, 23, 42)           # #0F172A (Slate 900)
    C_MUTED = RGBColor(100, 116, 139)       # #64748B (Slate 500)
    C_LIGHT = RGBColor(248, 250, 252)       # #F8FAFC
    C_WHITE = RGBColor(255, 255, 255)
    C_CARD_BG = RGBColor(241, 245, 249)     # Slate 100
    C_ACCENT_GREEN = RGBColor(22, 163, 74)  # Green 600

    blank_layout = prs.slide_layouts[6]

    def add_header(slide, title_text, category_text="SWASTH · MATERNAL & WOMEN'S HEALTH"):
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.1))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p0 = tf.paragraphs[0]
        p0.text = category_text.upper()
        p0.font.size = Pt(11)
        p0.font.bold = True
        p0.font.color.rgb = C_PRIMARY
        
        p1 = tf.add_paragraph()
        p1.text = title_text
        p1.font.size = Pt(24)
        p1.font.bold = True
        p1.font.color.rgb = C_DARK
        
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.65), Inches(11.733), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(226, 232, 240)
        line.line.color.rgb = RGBColor(226, 232, 240)

    # SLIDE 1: TITLE SLIDE
    slide1 = prs.slides.add_slide(blank_layout)
    bg = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_DARK
    bg.line.color.rgb = C_DARK

    title_box = slide1.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(11.3), Inches(4.8))
    tf1 = title_box.text_frame
    tf1.word_wrap = True
    
    p = tf1.paragraphs[0]
    p.text = "SWASTH (Maternal Health)"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY
    
    p2 = tf1.add_paragraph()
    p2.text = "AI-Powered Multilingual Long-Term Care Companion, Birth Clubs & Continuity Platform"
    p2.font.size = Pt(22)
    p2.font.color.rgb = C_WHITE
    p2.space_before = Pt(12)

    p3 = tf1.add_paragraph()
    p3.text = "Maternal & Women's Health Track  |  Primary User: Patient & Caregiver  |  Use Case: Long-Term Care Engagement"
    p3.font.size = Pt(13)
    p3.font.color.rgb = RGBColor(148, 163, 184)
    p3.space_before = Pt(18)

    p4 = tf1.add_paragraph()
    p4.text = "• Daily Garbh Sanskar Micro-Activities  • Hyperlocal Birth Clubs & Peer Support Circles\n• Multilingual Caregiver Sync (Hindi, Malayalam, Tamil)  • 24/7 Offline Access  • Zero Diagnostic Risk\n• Enterprise Open-Source Stack: Flutter + React + Python FastAPI + NVIDIA Nemotron / Llama 3 + PostgreSQL"
    p4.font.size = Pt(12)
    p4.font.color.rgb = RGBColor(203, 213, 225)
    p4.space_before = Pt(24)

    # SLIDE 2: THE PROBLEM
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "The Unseen Crisis: Why Long-Term Maternal Care Breaks Down")
    
    cards_data = [
        ("The Post-Discharge Cliff", 
         "Maternal and infant care requires 1,000+ days across 4 ANC checks, delivery, and 7 vaccine cohorts.\n\nYet, over 25-30% of mothers and infants drop out of scheduled milestones within months after hospital discharge.",
         C_PRIMARY),
        ("Administrative Overload & Isolation", 
         "Dropouts are rarely due to medical negligence—they stem from administrative friction and emotional isolation.\n\nCaregivers navigate fragile paper MCP cards, unclear clinic timings, and lack peer circles in their neighborhood.",
         C_SECONDARY),
        ("The 'Clinical App' Trap", 
         "Conventional health apps fail because they are cold and transactional. Patients open them only when an alert fires, leading to 80%+ abandonment in 30 days.\n\nWithout daily cultural rituals and peer community, long-term continuity collapses.",
         RGBColor(124, 58, 237))
    ]
    
    for i, (head, desc, col) in enumerate(cards_data):
        left = Inches(0.8 + i * 4.0)
        box = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(2.0), Inches(3.7), Inches(4.8))
        box.fill.solid()
        box.fill.fore_color.rgb = C_CARD_BG
        box.line.color.rgb = col
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.3)
        tf.margin_left = tf.margin_right = Inches(0.3)
        
        ph = tf.paragraphs[0]
        ph.text = head
        ph.font.size = Pt(18)
        ph.font.bold = True
        ph.font.color.rgb = col
        
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(13)
        pd.font.color.rgb = C_DARK
        pd.space_before = Pt(14)

    # SLIDE 3: THE SOLUTION - HOOK & BRIDGE
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "The Solution: Swasth's 'Daily Hook & Care Bridge' Engine")

    box_hook = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
    box_hook.fill.solid()
    box_hook.fill.fore_color.rgb = RGBColor(254, 242, 242)
    box_hook.line.color.rgb = C_PRIMARY
    box_hook.line.width = Pt(2)
    tf_hook = box_hook.text_frame
    tf_hook.word_wrap = True
    tf_hook.margin_left = tf_hook.margin_right = tf_hook.margin_top = Inches(0.3)
    
    p = tf_hook.paragraphs[0]
    p.text = "THE DAILY ENGAGEMENT HOOK"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY
    
    p = tf_hook.add_paragraph()
    p.text = "Garbh Sanskar Rituals & Hyperlocal Birth Clubs"
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = C_DARK
    p.space_before = Pt(4)

    points_hook = [
        "Daily 4Q Holistic Activities: 12 daily micro-habits across IQ, EQ (emotional), PQ (prenatal mobility), and SQ (spiritual calming).",
        "Hyperlocal Birth Clubs: Cohort circles based on expected delivery month and neighborhood/PHC radius for peer sharing and empathy.",
        "Garbh Samvaad & Raga Music: Soothing maternal lullabies, positive affirmations, and partner speaking sessions.",
        "Guardrail Moderation: NeMo AI guards peer questions against misinformation, instantly directing clinical symptoms to doctors."
    ]
    for pt in points_hook:
        p = tf_hook.add_paragraph()
        p.text = "• " + pt
        p.font.size = Pt(11.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(8)

    box_bridge = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(2.0), Inches(5.7), Inches(4.8))
    box_bridge.fill.solid()
    box_bridge.fill.fore_color.rgb = RGBColor(240, 249, 255)
    box_bridge.line.color.rgb = C_SECONDARY
    box_bridge.line.width = Pt(2)
    tf_bridge = box_bridge.text_frame
    tf_bridge.word_wrap = True
    tf_bridge.margin_left = tf_bridge.margin_right = tf_bridge.margin_top = Inches(0.3)

    p = tf_bridge.paragraphs[0]
    p.text = "THE HEALTHCARE CONTINUITY BRIDGE"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_SECONDARY

    p = tf_bridge.add_paragraph()
    p.text = "Vaccination, Diet & Caregiver Continuity"
    p.font.size = Pt(17)
    p.font.bold = True
    p.font.color.rgb = C_DARK
    p.space_before = Pt(4)

    points_bridge = [
        "Smart Vaccine Tracker: Automates 6, 10, 14-week and 9-month immunization schedules with vernacular voice reminders.",
        "Stage Weaning & Diet Plans: Doctor-approved month-by-month infant food charts and maternal lactation nutrition.",
        "Family & Caregiver Sync: Syncs reminders with partners, grandmothers, and local ASHA/ANMs via WhatsApp and voice.",
        "24/7 Offline Sync: Functions completely without internet, syncing encrypted data when signal resumes.",
        "Strict Non-Clinical Boundary: 0% diagnostic advice. Symptoms immediately route to primary health center emergency channels."
    ]
    for pt in points_bridge:
        p = tf_bridge.add_paragraph()
        p.text = "• " + pt
        p.font.size = Pt(11.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(8)

    # SLIDE 4: THE 6 CORE PILLARS
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Full Alignment with SWASTH's 6 Core Pillars")

    pillars = [
        ("01. Care Journey Companion", "Stage-Aware Guidance",
         ["Trimester checklists & delivery bag planner", "Weekly baby development & fruit-size analogies", "42-day postpartum lochia recovery tracking"],
         "Target: 95% Milestone Awareness"),
        ("02. Long-Term Care Engagement", "Sustained Multi-Year Adherence",
         ["Daily 4Q Garbh Sanskar routines (IQ, EQ, PQ, SQ)", "Milestone streak badges & celebration rewards", "Hyperlocal Birth Clubs for peer empathy circles"],
         "Target: >80% 90-Day Retention"),
        ("03. Family & Caregiver Support", "Caregiver Synchronization",
         ["Husband prep checklists & appointment sync", "Vernacular voice nudges for family elders", "Shared emergency transport & blood donor plans"],
         "Target: 3x Caregiver Participation"),
        ("04. Healthcare Navigation", "Clinic & Facility Guidance",
         ["Ultrasound scan window alerts & fasting steps", "PHC-to-District referral handoff documentation", "Pre-filled check-in questions for doctor visits"],
         "Target: 40% Shorter Clinic Wait Time"),
        ("05. Financial & Admin Support", "Maternity Scheme Navigator",
         ["PMMVY (₹5,000 grant) document tracker", "JSY & JSSK cash transfer entitlement guides", "Direct digital integration with ABHA ID"],
         "Target: 100% Scheme Realization"),
        ("06. Patient Journey Progress", "Longitudinal Milestone View",
         ["Digital MCP card sync with cloud backup", "Visual tracking of 7 infant vaccine cohorts", "Verified doctor visit summary archive"],
         "Target: Zero Lost Health Records")
    ]

    for idx, (p_tag, p_title, p_bullets, p_metric) in enumerate(pillars):
        row = idx // 3
        col = idx % 3
        left = Inches(0.8 + col * 4.0)
        top = Inches(1.9 + row * 2.6)
        box = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(3.75), Inches(2.45))
        box.fill.solid()
        box.fill.fore_color.rgb = C_CARD_BG
        box.line.color.rgb = C_PRIMARY if idx < 3 else C_SECONDARY
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.12)
        tf.margin_left = tf.margin_right = Inches(0.15)
        tf.margin_bottom = Inches(0.1)
        
        p0 = tf.paragraphs[0]
        p0.text = p_tag.upper()
        p0.font.size = Pt(9.5)
        p0.font.bold = True
        p0.font.color.rgb = C_PRIMARY
        
        p1 = tf.add_paragraph()
        p1.text = p_title
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = C_DARK
        p1.space_before = Pt(1)
        
        for b in p_bullets:
            pb = tf.add_paragraph()
            pb.text = "✓ " + b
            pb.font.size = Pt(10)
            pb.font.color.rgb = RGBColor(51, 65, 85)
            pb.space_before = Pt(2)
            
        pm = tf.add_paragraph()
        pm.text = "📊 " + p_metric
        pm.font.size = Pt(9.5)
        pm.font.bold = True
        pm.font.color.rgb = C_PRIMARY
        pm.space_before = Pt(4)

    # SLIDE 5: APP DESIGN SHOWCASE
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "App Design Showcase: Daily Rituals, Tools & Birth Clubs")

    img_preview_path = "/Volumes/DiskD/HACKATHONS/Swasth/assets/app_design/swasth_app_design_preview.png"
    if os.path.exists(img_preview_path):
        slide5.shapes.add_picture(img_preview_path, Inches(0.8), Inches(1.85), Inches(7.8), Inches(5.0))

    box_ux = slide5.shapes.add_textbox(Inches(8.8), Inches(1.85), Inches(3.8), Inches(5.0))
    tf_ux = box_ux.text_frame
    tf_ux.word_wrap = True
    tf_ux.margin_left = tf_ux.margin_top = tf_ux.margin_right = 0
    
    p = tf_ux.paragraphs[0]
    p.text = "CORE SCREENS SHOWCASED"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY
    
    features = [
        ("1. Daily 4Q Schedule", "12 daily activities across IQ, EQ, PQ, and SQ with day-by-day progress tracking."),
        ("2. Smart Tools & Scanner", "Quick access to MCP card scanner, water tracker, steps walk, and scheme aids."),
        ("3. Hyperlocal Birth Clubs", "Connecting mothers in nearby PHC circles with AI guardrail-moderated peer sharing."),
        ("4. Vaccine & Diet Tracker", "Upcoming 6/10/14-week vaccine cohorts and doctor-approved stage weaning meals.")
    ]
    for ft, fd in features:
        p = tf_ux.add_paragraph()
        p.text = ft
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(10)
        
        p = tf_ux.add_paragraph()
        p.text = fd
        p.font.size = Pt(11)
        p.font.color.rgb = C_MUTED
        p.space_before = Pt(2)

    # SLIDE 6: ENTERPRISE ARCHITECTURE
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Enterprise Open-Source Architecture: Speed & Reliability")

    tech_blocks = [
        ("Client Layer", "Flutter & Web",
         ["Flutter 3.x native (60–120 FPS)", "Hive local DB for 24/7 offline cache", "Lightweight React clinic backup portal", "<18 MB APK size for rural devices"],
         "⚡ <50ms UI Latency", C_PRIMARY),
        ("Unified Backend", "Python FastAPI",
         ["Async Starlette pipeline (high scale)", "WebSockets for live community sync", "ABHA Gateway tokenized login", "Background Celery immunization crons"],
         "⚡ 10k+ req/sec Throughput", C_SECONDARY),
        ("Agent Engine", "vLLM + LangGraph",
         ["Nemotron & Llama 3 on private cloud", "PagedAttention memory optimization", "LangGraph strict non-clinical paths", "Zero GPU lock-in (standard Linux)"],
         "⚡ <90ms First-Token Time", RGBColor(124, 58, 237)),
        ("Speech & NLP", "Indic NLP AI",
         ["Whisper STT fine-tuned on Indic accents", "AI4Bharat IndicTrans2 (10+ languages)", "Bhashini vernacular TTS voice notes", "15-sec localized WhatsApp audio nudges"],
         "⚡ 10+ Indic Languages", RGBColor(13, 148, 136)),
        ("Data & Audit", "PostgreSQL",
         ["Row-Level Security (RLS) tenant lock", "Immutable append-only audit ledger", "pgvector semantic guidance search", "DPDP Act 2023 & HIPAA compliance"],
         "⚡ 100% Cryptographic Log", RGBColor(217, 119, 6))
    ]

    for i, (layer_tag, layer_name, bullets, metric, col) in enumerate(tech_blocks):
        left = Inches(0.8 + i * 2.4)
        box = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.9), Inches(2.25), Inches(5.1))
        box.fill.solid()
        box.fill.fore_color.rgb = C_CARD_BG
        box.line.color.rgb = col
        box.line.width = Pt(2)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.15)
        tf.margin_left = tf.margin_right = Inches(0.15)
        tf.margin_bottom = Inches(0.1)
        
        p0 = tf.paragraphs[0]
        p0.text = layer_tag.upper()
        p0.font.size = Pt(9.5)
        p0.font.bold = True
        p0.font.color.rgb = col
        
        p1 = tf.add_paragraph()
        p1.text = layer_name
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = C_DARK
        p1.space_before = Pt(1)
        
        for b in bullets:
            pb = tf.add_paragraph()
            pb.text = "• " + b
            pb.font.size = Pt(10)
            pb.font.color.rgb = RGBColor(51, 65, 85)
            pb.space_before = Pt(6)
            
        pm = tf.add_paragraph()
        pm.text = metric
        pm.font.size = Pt(9.5)
        pm.font.bold = True
        pm.font.color.rgb = col
        pm.space_before = Pt(14)

    # SLIDE 7: SAFETY
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Clinical Safety: Zero Diagnostic Risk & Guardrail Moderation")

    safety_cols = [
        ("Strict Non-Clinical Boundary",
         "• 0% Diagnostic Advice: Never provides medical risk scores, diagnostic labels, or lab test interpretations.\n• 0% Treatment Prescriptions: Never suggests medications, dosage changes, or clinical therapies.\n• Non-Clinical Scope: Limited strictly to administrative scheduling, nutrition reminders, non-strenuous lifestyle, and welfare schemes.",
         C_PRIMARY),
        ("NVIDIA NeMo Guardrails",
         "• Real-time semantic guardrails run directly in memory on every user prompt and response.\n• Symptom Interception: If a mother mentions clinical symptoms (e.g. bleeding, severe pain, high fever), NeMo immediately halts automation.\n• Safe Direct Protocol: Instantly delivers emergency clinic helpline numbers and maps nearest PHC.",
         C_SECONDARY),
        ("Human-in-the-Loop & Auditability",
         "• Clinician & ANM Oversight: Frontline staff review milestone rosters and due-lists with 1-click approvals.\n• Append-Only Audit Trail: Every AI-generated recommendation, user interaction, and clinic override is permanently logged in PostgreSQL.\n• Full Patient & Family Consent: Caregiver synchronization requires explicit verification.",
         C_ACCENT_GREEN)
    ]

    for i, (title, text, col) in enumerate(safety_cols):
        left = Inches(0.8 + i * 4.0)
        box = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(2.0), Inches(3.7), Inches(4.8))
        box.fill.solid()
        box.fill.fore_color.rgb = C_CARD_BG
        box.line.color.rgb = col
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.3)
        tf.margin_left = tf.margin_right = Inches(0.25)
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = col
        
        p = tf.add_paragraph()
        p.text = text
        p.font.size = Pt(11.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(12)

    # SLIDE 8: IMPACT & ROADMAP
    slide8 = prs.slides.add_slide(blank_layout)
    add_header(slide8, "Measurable Operational Outcomes within 60–90 Days")

    metrics = [
        ("85% Drop", "Frontline Admin Burden", "Reduces ASHA/ANM time spent tallying paper due-lists from 10 hrs to <1.5 hrs/week."),
        ("90%+ Retention", "Immunization Milestones", "Reduces 6 to 14-week infant vaccine dropouts through localized caregiver voice alerts."),
        ("4x Higher", "Caregiver & Peer Activity", "Birth Clubs & WhatsApp synchronization drive proactive family and community involvement."),
        ("100% Safe", "Regulatory Compliance", "Full audit logging and zero unreviewed clinical recommendations.")
    ]

    for i, (val, lbl, sub) in enumerate(metrics):
        left = Inches(0.8 + i * 3.0)
        box = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(2.0), Inches(2.8), Inches(2.2))
        box.fill.solid()
        box.fill.fore_color.rgb = C_CARD_BG
        box.line.color.rgb = C_PRIMARY
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.2)
        tf.margin_left = tf.margin_right = Inches(0.15)
        
        p = tf.paragraphs[0]
        p.text = val
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = C_PRIMARY
        
        p = tf.add_paragraph()
        p.text = lbl
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(4)
        
        p = tf.add_paragraph()
        p.text = sub
        p.font.size = Pt(10)
        p.font.color.rgb = C_MUTED
        p.space_before = Pt(4)

    box_road = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(4.5), Inches(11.733), Inches(2.3))
    box_road.fill.solid()
    box_road.fill.fore_color.rgb = RGBColor(240, 253, 244)
    box_road.line.color.rgb = C_ACCENT_GREEN
    box_road.line.width = Pt(1.5)
    tf_road = box_road.text_frame
    tf_road.word_wrap = True
    tf_road.margin_top = Inches(0.2)
    tf_road.margin_left = tf_road.margin_right = Inches(0.3)
    
    p = tf_road.paragraphs[0]
    p.text = "EXECUTION & ROLLOUT ROADMAP"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_ACCENT_GREEN
    
    phases = [
        "Day 1–30 (Prototype & Guardrail Audit): Deploy Flutter mobile client + Python backend; benchmark open-source Nemotron on vLLM; validate NeMo non-clinical guardrails.",
        "Day 31–60 (Pilot with 5 PHCs & Birth Clubs): Onboard 500 expectant mothers & caregivers; launch 5 localized Birth Clubs; track daily Garbh Sanskar engagement and vaccine adherence.",
        "Day 61–90 (Longitudinal Scaling & Portal Integration): Synchronize with state HMIS/RCH registers; export verified administrative audit reports; expand vernacular voice models into 10+ dialects."
    ]
    for ph in phases:
        p = tf_road.add_paragraph()
        p.text = "➔ " + ph
        p.font.size = Pt(11)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(4)

    output_path = "/Volumes/DiskD/HACKATHONS/Swasth/deck/Swasth_Presentation.pptx"
    prs.save(output_path)
    print(f"Updated PowerPoint Presentation successfully created at {output_path}")

if __name__ == "__main__":
    create_presentation()
