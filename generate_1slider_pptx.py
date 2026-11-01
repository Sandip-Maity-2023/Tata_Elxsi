import os
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_1slider():
    prs = pptx.Presentation()
    # 16:9 widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Palette
    C_BG_PAGE = RGBColor(241, 245, 249)      # Slate 100
    C_NAVY_DARK = RGBColor(10, 25, 47)       # Deep Corporate Navy #0A192F
    C_NAVY_HEADER = RGBColor(15, 30, 60)     # Navy #0F1E3C
    C_BLUE_ACCENT = RGBColor(0, 114, 206)    # Tata Blue #0072CE
    C_CYAN_TECH = RGBColor(14, 165, 233)     # Sky Cyan #0EA5E9
    C_WHITE = RGBColor(255, 255, 255)
    C_TEXT_DARK = RGBColor(15, 23, 42)       # Slate 900
    C_TEXT_MUTED = RGBColor(71, 85, 105)     # Slate 600
    C_BORDER = RGBColor(203, 213, 225)       # Slate 300
    
    # Severity Colors
    C_CRIT = RGBColor(220, 38, 38)          # Red #DC2626
    C_WARN = RGBColor(234, 88, 12)          # Orange #EA580C
    C_CAUT = RGBColor(217, 119, 6)          # Amber #D97706
    C_OK = RGBColor(16, 185, 129)           # Emerald #10B981

    # Base Slide Background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG_PAGE
    bg.line.fill.background()

    # -------------------------------------------------------------
    # TOP HEADER BANNER (Navy)
    # -------------------------------------------------------------
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.02))
    header.fill.solid()
    header.fill.fore_color.rgb = C_NAVY_DARK
    header.line.fill.background()

    # Header Text
    tf_h = header.text_frame
    tf_h.word_wrap = True
    tf_h.margin_left = Inches(0.4)
    tf_h.margin_top = Inches(0.1)
    tf_h.margin_right = Inches(0.4)
    tf_h.margin_bottom = Inches(0.05)

    p0 = tf_h.paragraphs[0]
    p0.text = "TATA ELXSI | TELIPORT SEASON 4 — CASE STUDY 1: PREDICTIVE MAINTENANCE FOR CONNECTED FLEET"
    p0.font.size = Pt(9.5)
    p0.font.bold = True
    p0.font.color.rgb = C_CYAN_TECH

    p1 = tf_h.add_paragraph()
    p1.text = "AEROLINK-PHM: Autonomous Closed-Loop Anomaly Detection, RUL Prognostics & Rule-Based Work Order Orchestration"
    p1.font.size = Pt(15.5)
    p1.font.bold = True
    p1.font.color.rgb = C_WHITE

    p2 = tf_h.add_paragraph()
    p2.text = "NASA C-MAPSS Benchmark (FD001–FD004)  •  100% Non-LLM Deterministic Triage  •  Multi-Regime Normalization  •  End-to-End MLOps Closed-Loop"
    p2.font.size = Pt(8.5)
    p2.font.color.rgb = RGBColor(203, 213, 225)

    # -------------------------------------------------------------
    # TOP PIPELINE FLOW STRIP (0.32 in high)
    # -------------------------------------------------------------
    flow_y = Inches(1.08)
    flow_h = Inches(0.32)
    pipeline_steps = [
        ("1. INGESTION & REGIMES", "21 Sensors + 3 Settings (FD001-FD004)"),
        ("2. DATA ENGINEERING", "K-Means Regime Clust. + EWMA Denoise"),
        ("3. ML PROGNOSTICS", "IsoForest Detect + LightGBM RUL (Asymm. Loss)"),
        ("4. NON-LLM TRIAGE AGENT", "4-Tier State Machine + Auto Work Order"),
        ("5. MLOps & RETRAINING", "MLflow Track + KS Drift Retraining Loop")
    ]
    step_w = Inches(2.45)
    gap = Inches(0.09)
    start_x = Inches(0.35)

    for i, (title, sub) in enumerate(pipeline_steps):
        x = start_x + i * (step_w + gap)
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, flow_y, step_w, flow_h)
        box.fill.solid()
        box.fill.fore_color.rgb = C_NAVY_HEADER
        box.line.color.rgb = C_BLUE_ACCENT
        box.line.width = Pt(1)

        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        tf.margin_left = Inches(0.06)
        tf.margin_right = Inches(0.06)
        
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = f"{title}: "
        p.font.size = Pt(7.5)
        p.font.bold = True
        p.font.color.rgb = C_CYAN_TECH
        
        run = p.add_run()
        run.text = sub
        run.font.bold = False
        run.font.color.rgb = C_WHITE
        run.font.size = Pt(7)

    # -------------------------------------------------------------
    # 4 MAIN CONTENT COLUMNS
    # Layout dimensions
    # -------------------------------------------------------------
    col_y = Inches(1.46)
    col_h = Inches(5.35)
    col_w = Inches(3.02)
    col_gap = Inches(0.18)
    col_x_start = Inches(0.35)

    # Helper function to create card
    def create_card(slide, x, y, w, h, header_title, weight_badge, accent_color):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        card.fill.solid()
        card.fill.fore_color.rgb = C_WHITE
        card.line.color.rgb = C_BORDER
        card.line.width = Pt(1)

        # Header bar on top of card
        top_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, Inches(0.42))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = accent_color
        top_bar.line.fill.background()

        tf = top_bar.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.06)
        tf.margin_left = Inches(0.12)
        tf.margin_right = Inches(0.12)

        p = tf.paragraphs[0]
        p.text = header_title.upper()
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = C_WHITE

        # Badge text right aligned
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.RIGHT
        p2.text = f"EVAL WEIGHT: {weight_badge}"
        p2.font.size = Pt(6.5)
        p2.font.bold = True
        p2.font.color.rgb = RGBColor(224, 242, 254)

        return card

    # =============================================================
    # COLUMN 1: DATA ENGINEERING & REGIME NORMALIZATION
    # =============================================================
    col1_x = col_x_start + 0 * (col_w + col_gap)
    create_card(slide, col1_x, col_y, col_w, col_h, "1. Data Engineering & Physics", "15% + 10%", C_NAVY_HEADER)

    tb1 = slide.shapes.add_textbox(col1_x + Inches(0.08), col_y + Inches(0.46), col_w - Inches(0.16), col_h - Inches(0.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = tf1.margin_right = tf1.margin_top = tf1.margin_bottom = 0

    def add_sec(tf, heading, items):
        p_head = tf.add_paragraph()
        p_head.text = heading
        p_head.font.size = Pt(8.2)
        p_head.font.bold = True
        p_head.font.color.rgb = C_BLUE_ACCENT
        p_head.space_before = Pt(4)
        p_head.space_after = Pt(1)

        for bullet, text in items:
            p = tf.add_paragraph()
            p.text = f"• {bullet}: "
            p.font.size = Pt(7.0)
            p.font.bold = True
            p.font.color.rgb = C_TEXT_DARK
            p.space_after = Pt(1.5)
            
            run = p.add_run()
            run.text = text
            run.font.bold = False
            run.font.color.rgb = C_TEXT_MUTED

    add_sec(tf1, "TURBOFAN SENSOR ANALYSIS & PRUNING", [
        ("NASA C-MAPSS Physics", "21 continuous sensors + 3 flight operating settings across run-to-failure engine trajectories."),
        ("Zero-Variance Pruning", "Sensors 1, 5, 10, 16, 18, 19 drop constant variance (stdev=0) across flights to avoid collinear noise.")
    ])

    add_sec(tf1, "MULTI-REGIME FLIGHT NORMALIZATION", [
        ("Operating Conditions", "FD001/003 operate at sea-level; FD002/004 span 6 flight regimes (Alt: 0-42k ft, Mach: 0-0.84, TRA: 20-100)."),
        ("K-Means Regime Clust.", "Cluster operational settings into 6 operational centroids to segment distinct flight physics."),
        ("Regime Z-Score Scaling", "Standardize each sensor per regime k: Z = (X - μ_k) / σ_k. Removes environmental baseline swings while preserving true wear drift.")
    ])

    add_sec(tf1, "DENOISING & FEATURE ENGINEERING", [
        ("Noise Attenuation", "Savitzky-Golay / EWMA filter (α=0.2) removes high-frequency turbulence without lag in degradation slope."),
        ("Rolling-Window Features", "Extract rolling mean, stdev, min, max, peak-to-peak over 15 & 30 cycles to capture degradation velocity."),
        ("Degradation Trend (d/dt)", "Linear trend slope over sliding window isolates monotonic wear vs nominal cyclical fluctuations."),
        ("Thermodynamic Ratios", "Compute physics interactions: T30/P30 (combustor pressure-heat ratio), Ps30/Nc (HPC pressure/speed efficiency).")
    ])

    # =============================================================
    # COLUMN 2: MACHINE LEARNING & PROGNOSTICS (RUL)
    # =============================================================
    col2_x = col_x_start + 1 * (col_w + col_gap)
    create_card(slide, col2_x, col_y, col_w, col_h, "2. ML Prognostics Engine", "20%", C_NAVY_HEADER)

    tb2 = slide.shapes.add_textbox(col2_x + Inches(0.08), col_y + Inches(0.46), col_w - Inches(0.16), col_h - Inches(0.5))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = tf2.margin_right = tf2.margin_top = tf2.margin_bottom = 0

    add_sec(tf2, "TWO-STAGE HYBRID ARCHITECTURE", [
        ("Stage 1 - Anomaly Inception", "Unsupervised Isolation Forest + Mahalanobis Distance on nominal baseline to detect first fault inflection point (t_fault)."),
        ("Stage 2 - RUL Estimation", "Active only once anomaly threshold breached, avoiding artificial early degradation penalties.")
    ])

    add_sec(tf2, "PIECEWISE-LINEAR RUL TARGET MODEL", [
        ("Physics-Informed RUL Cap", "Early engine life shows negligible physical wear. Ground-truth RUL is clipped at RUL_max = 125 cycles."),
        ("Degradation Progression", "Linear descent towards 0 at failure cycle. Prevents regressor from fitting flat initial operational wear noise.")
    ])

    add_sec(tf2, "ALGORITHM SELECTION & REASONING", [
        ("Primary - LightGBM Regressor", "High inference speed (<5ms), leaf-wise tree splitting, native handling of correlated rolling features, robust to outliers."),
        ("Secondary - Bi-LSTM / CNN", "Deep temporal alternative for complex temporal sequence modeling in FD004 benchmark comparison."),
        ("Imbalance Mitigation", "Time-decay sample re-weighting emphasizes critical low-RUL cycles (RUL < 30) over abundant healthy flight cycles.")
    ])

    add_sec(tf2, "NASA PHM'08 ASYMMETRIC LOSS FUNCTION", [
        ("Operational Risk Reality", "In aviation/fleet maintenance, late predictions (d > 0) risk catastrophic in-flight failure, while early alerts (d < 0) cost routine inspections."),
        ("Asymmetric Penalty Metric", "S = Σ [exp(-d/13) - 1] for d < 0 vs Σ [exp(d/10) - 1] for d ≥ 0. Heavily penalizes dangerous underestimations of wear."),
        ("Explainable AI (TreeSHAP)", "Local Shapley values attribute top degrading sensors (e.g., T30, Ps30) for transparent root-cause verification.")
    ])

    # =============================================================
    # COLUMN 3: NON-LLM TRIAGE AGENT & WORK ORDERS
    # =============================================================
    col3_x = col_x_start + 2 * (col_w + col_gap)
    create_card(slide, col3_x, col_y, col_w, col_h, "3. Non-LLM Triage Agent", "15%", C_NAVY_HEADER)

    tb3 = slide.shapes.add_textbox(col3_x + Inches(0.08), col_y + Inches(0.46), col_w - Inches(0.16), col_h - Inches(0.5))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    tf3.margin_left = tf3.margin_right = tf3.margin_top = tf3.margin_bottom = 0

    add_sec(tf3, "DETERMINISTIC RULE ENGINE (NO LLM)", [
        ("Ground Rule Compliance", "100% Deterministic Python state machine. Zero hallucination risk, ultra-low memory, sub-millisecond edge execution."),
        ("Closed-Loop Triage", "Consumes Model RUL + Anomaly Score + Degradation Velocity to trigger automated actions without human intervention.")
    ])

    add_sec(tf3, "4-TIER SEVERITY ACTION MATRIX", [
        ("🔴 CRITICAL (RUL ≤ 15 cycles)", "Immediate Flight Grounding / Lockout. P1 Emergency Work Order auto-dispatched. Pre-allocates HPC Blade Kit SKU-402."),
        ("🟠 WARNING (15 < RUL ≤ 45)", "P2 Base Turnaround Maintenance (<48 hrs). Automated maintenance dock routing & technician scheduling."),
        ("🟡 CAUTION (45 < RUL ≤ 80)", "P3 High-Frequency Polling (5 Hz). Unit flagged in fleet health watchlist; condition tracking enabled."),
        ("🟢 NOMINAL (RUL > 80)", "Standard line maintenance clearance; nominal flight duty cycle assigned.")
    ])

    add_sec(tf3, "AUTOMATED CMMS WORK ORDER SCHEMA", [
        ("Automated Payload", "Generates compliant JSON payload to enterprise ERP / CMMS (SAP PM / IBM Maximo)."),
        ("Key Schema Fields", "WO_ID, Unit_ID, Priority, Pred_RUL ± 95% CI, Degradation_Rate, Root_Cause_Sensor, Action_Items, Parts_SKU.")
    ])

    add_sec(tf3, "HUMAN-IN-THE-LOOP EXCEPTION ESCALATION", [
        ("Exception Guardrails", "Human Reliability Engineer steps in ONLY for: (1) Prediction uncertainty variance > 25%, (2) Conflicting sensor alarms, (3) Anomaly without RUL drop."),
        ("Engineer Triage Card", "Renders one-click telemetry summary, residual plot, and SHAP attribution for rapid override/approval.")
    ])

    # =============================================================
    # COLUMN 4: MLOPS & DATASET SCALABILITY
    # =============================================================
    col4_x = col_x_start + 3 * (col_w + col_gap)
    create_card(slide, col4_x, col_y, col_w, col_h, "4. MLOps & Dataset Coverage", "15% + 10%", C_NAVY_HEADER)

    tb4 = slide.shapes.add_textbox(col4_x + Inches(0.08), col_y + Inches(0.46), col_w - Inches(0.16), col_h - Inches(0.5))
    tf4 = tb4.text_frame
    tf4.word_wrap = True
    tf4.margin_left = tf4.margin_right = tf4.margin_top = tf4.margin_bottom = 0

    add_sec(tf4, "MLOps LIFECYCLE & MONITORING", [
        ("MLflow Tracking & Registry", "Logs hyperparameters, RMSE, PHM scores, and model binaries. Production staging tag gated by validation score."),
        ("Data & Concept Drift Detect", "Kolmogorov-Smirnov test on rolling sensor inputs + Population Stability Index (PSI > 0.25) to detect sensor calibration drift."),
        ("Automated Retraining Loop", "Triggered when test RMSE degrades >12% or when 10 new run-to-failure / teardown inspection records are ingested."),
        ("Streamlit Fleet Dashboard", "Live fleet health map, unit RUL countdowns, sensor degradation graphs, and pending work order triage queue.")
    ])

    add_sec(tf4, "NASA C-MAPSS DATASET SCALABILITY (FD001–FD004)", [
        ("FD001 (1 Cond, 1 Fault - HPC)", "Baseline validation. Single regime, simple degradation dynamics. Target RMSE < 14.5, Score < 350."),
        ("FD002 (6 Cond, 1 Fault - HPC)", "High operating variability. Handled via 6-centroid K-Means regime normalization. Target RMSE < 20.5."),
        ("FD003 (1 Cond, 2 Faults - HPC/Fan)", "Compound fault modes. Incorporates Random Forest fault-mode classifier before branch RUL regression."),
        ("FD004 (6 Cond, 2 Faults - HPC/Fan)", "Full industrial complexity. Hierarchical pipeline: Regime Clust. ➔ Multi-Fault Isolation ➔ Ensemble RUL Regress.")
    ])

    add_sec(tf4, "BUSINESS IMPACT & ROI (FLEET VALUE)", [
        ("35% Downtime Reduction", "Eliminates catastrophic in-flight failure via proactive base-dock servicing."),
        ("22% Inventory Savings", "Automated pre-allocation of spares based on 45-cycle advance notice."),
        ("98% Autonomous Execution", "Human engineers intervene strictly on exceptions (<2% of total events).")
    ])

    # -------------------------------------------------------------
    # BOTTOM FOOTER BAR (0.42 in high)
    # -------------------------------------------------------------
    footer_y = Inches(6.88)
    footer_h = Inches(0.48)
    footer = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, footer_y, Inches(13.333), footer_h)
    footer.fill.solid()
    footer.fill.fore_color.rgb = C_NAVY_DARK
    footer.line.fill.background()

    tf_f = footer.text_frame
    tf_f.word_wrap = True
    tf_f.margin_top = Inches(0.08)
    tf_f.margin_bottom = Inches(0.04)
    tf_f.margin_left = Inches(0.4)
    tf_f.margin_right = Inches(0.4)

    pf = tf_f.paragraphs[0]
    pf.alignment = PP_ALIGN.LEFT
    pf.text = "ROUND 1 DELIVERABLES COVERAGE: "
    pf.font.size = Pt(8.0)
    pf.font.bold = True
    pf.font.color.rgb = C_CYAN_TECH

    deliverables = [
        "[✓] End-to-End Architecture Diagram (Data → Features → Model → Agent → Work Order)",
        "[✓] Data Cleaning & Rolling Feature Engineering",
        "[✓] Dual ML Anomaly Detection & RUL Models with Asymmetric Loss Reasoning",
        "[✓] Deterministic Non-LLM Triage Agent Logic & Schema",
        "[✓] MLOps Retraining & Multi-Dataset Scalability (FD001-FD004)"
    ]
    
    run_f = pf.add_run()
    run_f.text = "   " + "   •   ".join(deliverables)
    run_f.font.size = Pt(7.2)
    run_f.font.bold = False
    run_f.font.color.rgb = C_WHITE

    # Output path
    out_path = r"E:\Tata_Elxsi\Tata_Elxsi_TELIPORT_Predictive_Maintenance_1Slider.pptx"
    prs.save(out_path)
    print(f"Presentation saved successfully to: {out_path}")

if __name__ == "__main__":
    create_1slider()

