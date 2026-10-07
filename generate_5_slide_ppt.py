import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_5_slide_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    C_PRIMARY = RGBColor(233, 59, 103)      # Rose Pink
    C_SECONDARY = RGBColor(99, 102, 241)    # Indigo
    C_DARK = RGBColor(15, 23, 42)           # Slate 900
    C_MUTED = RGBColor(100, 116, 139)       # Slate 500
    C_LIGHT_BG = RGBColor(248, 250, 252)    # Slate 50
    C_BORDER = RGBColor(226, 232, 240)      # Slate 200
    C_ACCENT_GREEN = RGBColor(16, 185, 129) # Mint Green

    blank_layout = prs.slide_layouts[6]

    def add_slide_header(slide, title_text, category_text="SWASTH · APP DESIGN & USER JOURNEY"):
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.7), Inches(1.1))
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
        
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.6), Inches(11.733), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = C_BORDER
        line.line.color.rgb = C_BORDER

    # =========================================================================
    # SLIDE 1: APP IDENTITY & HERO SHOWCASE
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = C_DARK
    bg1.line.color.rgb = C_DARK

    box1 = slide1.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(6.5), Inches(5.2))
    tf1 = box1.text_frame
    tf1.word_wrap = True
    
    p = tf1.paragraphs[0]
    p.text = "SWASTH (Maternal Health)"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY
    
    p = tf1.add_paragraph()
    p.text = "Track Every Moment with Care: From Pregnancy to the First 1,000 Days"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.space_before = Pt(8)

    p = tf1.add_paragraph()
    p.text = "Maternal & Women's Health Track  |  Patient & Caregiver Long-Term Engagement"
    p.font.size = Pt(12)
    p.font.color.rgb = RGBColor(148, 163, 184)
    p.space_before = Pt(14)

    p = tf1.add_paragraph()
    p.text = "• Modern 3D Stage Progress: Week-by-week fetal growth, trimesters & investigation schedulers\n• Daily Garbh Sanskar Rituals: IQ, EQ, PQ, SQ micro-activities & soothing raga lullabies\n• Hyperlocal Birth Clubs: Community peer circles with NeMo Guardrails safety moderation\n• Multilingual Caregiver Sync: Voice & WhatsApp alerts for partners & elders in native dialects\n• 24/7 Offline Sync: Zero data loss even without internet connectivity"
    p.font.size = Pt(12)
    p.font.color.rgb = RGBColor(226, 232, 240)
    p.space_before = Pt(20)

    hero_img = "/Volumes/DiskD/HACKATHONS/Swasth/website/assets/app_screens/pillar1_1_home_overview.png"
    if os.path.exists(hero_img):
        slide1.shapes.add_picture(hero_img, Inches(7.5), Inches(1.2), Inches(5.2), Inches(5.5))

    # =========================================================================
    # SLIDE 2: INFORMATION ARCHITECTURE & USER JOURNEY
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    add_slide_header(slide2, "App Information Architecture & Complete User Journey")

    tree_img = "/Volumes/DiskD/HACKATHONS/Swasth/website/assets/app_screens/pillar1_4_home_vitals.png"
    if os.path.exists(tree_img):
        slide2.shapes.add_picture(tree_img, Inches(0.8), Inches(1.8), Inches(5.4), Inches(5.0))

    intent_img = "/Volumes/DiskD/HACKATHONS/Swasth/website/assets/app_screens/pillar1_3_home_milestones.png"
    if os.path.exists(intent_img):
        slide2.shapes.add_picture(intent_img, Inches(6.4), Inches(1.8), Inches(3.6), Inches(5.0))

    box2 = slide2.shapes.add_textbox(Inches(10.2), Inches(1.8), Inches(2.4), Inches(5.0))
    tf2 = box2.text_frame
    tf2.word_wrap = True
    
    p = tf2.paragraphs[0]
    p.text = "CORE FLOW"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY

    flow_items = [
        ("Onboarding Intent", "Follow baby growth, learn tips, and plan weeks safely."),
        ("Home & 3D Stage", "Trimesters, baby size, daily 4Q rituals, and clinic slots."),
        ("Smart Tools", "Contraction timer, kick counter, MCP card photo scanner."),
        ("Birth Clubs", "Peer circles with NeMo Guardrails moderation."),
        ("Caregiver Hub", "Partner sync via WhatsApp & voice alerts in dialects.")
    ]
    for h, d in flow_items:
        p = tf2.add_paragraph()
        p.text = h
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(8)
        
        p = tf2.add_paragraph()
        p.text = d
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_MUTED
        p.space_before = Pt(1)

    # =========================================================================
    # SLIDE 3: MODERN DESIGN SYSTEM & ONBOARDING SUITE
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    add_slide_header(slide3, "Modern Design System: 3D Micro-Interactions & Due Date Setup")

    duedate_img = "/Volumes/DiskD/HACKATHONS/Swasth/website/assets/app_screens/pillar2_1_plans_toolkit.png"
    if os.path.exists(duedate_img):
        slide3.shapes.add_picture(duedate_img, Inches(0.8), Inches(1.8), Inches(4.5), Inches(5.0))

    growth_img = "/Volumes/DiskD/HACKATHONS/Swasth/website/assets/app_screens/pillar2_2_kick_counter.png"
    if os.path.exists(growth_img):
        slide3.shapes.add_picture(growth_img, Inches(5.5), Inches(1.8), Inches(4.5), Inches(5.0))

    box3 = slide3.shapes.add_textbox(Inches(10.2), Inches(1.8), Inches(2.4), Inches(5.0))
    tf3 = box3.text_frame
    tf3.word_wrap = True

    p = tf3.paragraphs[0]
    p.text = "DESIGN HIGHLIGHTS"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY

    design_specs = [
        ("Flexible Due Date Estimation", "Supports Gestational age, LMP date, Conception date, or manual due date."),
        ("Tactile 3D Visuals", "Apple/peach baby size comparisons & joyful floating heart milestones."),
        ("Soft Pastel Tones", "Warm rose (#E93B67) and lavender to eliminate clinical stress."),
        ("Caregiver First", "Simple large buttons designed for low digital literacy.")
    ]
    for h, d in design_specs:
        p = tf3.add_paragraph()
        p.text = h
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(8)
        
        p = tf3.add_paragraph()
        p.text = d
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_MUTED
        p.space_before = Pt(1)

    # =========================================================================
    # SLIDE 4: AI PREGNANCY ASSISTANT & 4 SIGNATURE MOBILE EXPERIENCES
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    add_slide_header(slide4, "AI Assistant Chat & Signature Experiences (Strictly Non-Clinical)")

    # Left: AI Assistant Chat Screen / Timed Circuit
    ai_chat_img = "/Volumes/DiskD/HACKATHONS/Swasth/website/assets/app_screens/pillar4_2_timed_circuit.png"
    if os.path.exists(ai_chat_img):
        slide4.shapes.add_picture(ai_chat_img, Inches(0.8), Inches(1.8), Inches(4.2), Inches(5.0))

    # Center: Signature Experiences & Nutrition
    app_preview = "/Volumes/DiskD/HACKATHONS/Swasth/website/assets/app_screens/pillar4_4_nutrition_recipes.png"
    if os.path.exists(app_preview):
        slide4.shapes.add_picture(app_preview, Inches(5.2), Inches(1.8), Inches(5.0), Inches(5.0))

    # Right: Feature details
    box4 = slide4.shapes.add_textbox(Inches(10.4), Inches(1.8), Inches(2.2), Inches(5.0))
    tf4 = box4.text_frame
    tf4.word_wrap = True

    p = tf4.paragraphs[0]
    p.text = "CORE EXPERIENCES"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY

    screens_list = [
        ("AI Companion Chat", "NeMo Guardrail verified non-medical chat: Nutrition, timeline & wellness."),
        ("3D Stage Journey", "Trimester rail, Apple-size fetal progress, ultrasound appointments."),
        ("4Q Garbh Sanskar", "Daily 12 micro-habits & calming raga lullabies."),
        ("Birth Clubs", "Hyperlocal mother circles with peer safety moderation.")
    ]
    for h, d in screens_list:
        p = tf4.add_paragraph()
        p.text = h
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(8)
        
        p = tf4.add_paragraph()
        p.text = d
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_MUTED
        p.space_before = Pt(1)

    # =========================================================================
    # SLIDE 5: ENTERPRISE ARCHITECTURE, DATA PROTECTION & IMPACT
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    add_slide_header(slide5, "Enterprise Architecture, Data Protection & 60-90 Day Impact")

    # Left: Settings & Data Protection image
    settings_img = "/Volumes/DiskD/HACKATHONS/Swasth/website/assets/app_screens/pillar5_2_abha_privacy.png"
    if os.path.exists(settings_img):
        slide5.shapes.add_picture(settings_img, Inches(0.8), Inches(1.8), Inches(3.8), Inches(5.0))

    col_data = [
        ("Unified Open-Source Stack",
         "• Flutter Mobile App: Native 60-120 FPS iOS & Android with encrypted offline sync.\n• React/TypeScript Web: Backup portal for care coordinators.\n• Python (FastAPI): Single unified REST & WebSockets backend.\n• NVIDIA Nemotron / Llama 3: Open-weights models served on vLLM with PagedAttention.\n• Indic NLP: Open-source Whisper & AI4Bharat IndicTrans2 for 10+ Indian languages.\n• Enterprise PostgreSQL: Row-level security & immutable audit trail.",
         C_PRIMARY),
        ("Safety & 60-90 Day Impact",
         "• 0% Diagnostic Advice: Strictly non-medical.\n• NVIDIA NeMo Guardrails: Intercepts symptoms and routes to PHC emergency.\n• Human-in-the-Loop: Clinicians review milestone rosters with 1-click approvals.\n• 85% Drop in Admin Burden: Cuts manual ASHA/ANM due-list compilation to <1.5 hrs/week.\n• 90%+ Vaccine Adherence: Reduces 6 to 14-week infant immunization dropouts.\n• 4x Higher Engagement: Birth Clubs & caregiver WhatsApp synchronization.",
         C_SECONDARY)
    ]

    for i, (head, body, col) in enumerate(col_data):
        left = Inches(4.8 + i * 4.0)
        box = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.8), Inches(3.8), Inches(5.0))
        box.fill.solid()
        box.fill.fore_color.rgb = C_LIGHT_BG
        box.line.color.rgb = col
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.22)
        tf.margin_left = tf.margin_right = Inches(0.22)
        
        p = tf.paragraphs[0]
        p.text = head
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = col
        
        p = tf.add_paragraph()
        p.text = body
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(8)

    output_path = "/Volumes/DiskD/HACKATHONS/Swasth/deck/Swasth_App_Design_5_Slides.pptx"
    prs.save(output_path)
    print(f"5-Slide App Design PowerPoint successfully updated at {output_path}")

if __name__ == "__main__":
    create_5_slide_presentation()
