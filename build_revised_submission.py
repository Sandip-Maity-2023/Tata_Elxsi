from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


OUTPUT = Path("Teliport_S4_Case1_AEROLINK-PHM_1Slide_Revised.pptx")

NAVY = RGBColor(7, 19, 38)
NAVY_2 = RGBColor(16, 39, 70)
BLUE = RGBColor(0, 114, 206)
CYAN = RGBColor(14, 165, 233)
TEAL = RGBColor(13, 148, 136)
BG = RGBColor(244, 247, 251)
WHITE = RGBColor(255, 255, 255)
INK = RGBColor(22, 36, 55)
MUTED = RGBColor(71, 85, 105)
BORDER = RGBColor(203, 213, 225)
GREEN = RGBColor(5, 150, 105)
AMBER = RGBColor(217, 119, 6)
ORANGE = RGBColor(234, 88, 12)
RED = RGBColor(220, 38, 38)


def fill(shape, color, line=None):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line


def text_box(slide, x, y, w, h, text, size=9, color=INK, bold=False,
             align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, font="Aptos", margin=0.03):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    paragraph = tf.paragraphs[0]
    paragraph.alignment = align
    paragraph.text = text
    paragraph.font.name = font
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold
    paragraph.font.color.rgb = color
    return box


def card(slide, x, y, w, h, title, accent=BLUE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    fill(shape, WHITE, BORDER)
    shape.line.width = Pt(0.8)
    strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.07))
    fill(strip, accent)
    text_box(slide, x + 0.12, y + 0.12, w - 0.24, 0.28, title, size=9.2, color=NAVY_2, bold=True)
    return shape


def body(slide, x, y, w, lines, size=7.1):
    offset = y
    for heading, description in lines:
        box = slide.shapes.add_textbox(Inches(x), Inches(offset), Inches(w), Inches(0.42))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.space_after = Pt(0)
        run = p.add_run()
        run.text = heading
        run.font.name = "Aptos"
        run.font.size = Pt(size)
        run.font.bold = True
        run.font.color.rgb = INK
        run = p.add_run()
        run.text = description
        run.font.name = "Aptos"
        run.font.size = Pt(size)
        run.font.color.rgb = MUTED
        offset += 0.44


def flow_node(slide, x, title, detail, accent):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.62), Inches(2.25), Inches(1.06))
    fill(shape, WHITE, BORDER)
    shape.line.width = Pt(0.8)
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1.62), Inches(2.25), Inches(0.12))
    fill(band, accent)
    text_box(slide, x + 0.11, 1.82, 2.03, 0.24, title, size=8.1, color=NAVY_2, bold=True, align=PP_ALIGN.CENTER)
    text_box(slide, x + 0.12, 2.09, 2.01, 0.43, detail, size=6.5, color=MUTED, align=PP_ALIGN.CENTER)


def arrow(slide, x):
    text_box(slide, x, 1.91, 0.32, 0.26, "›", size=24, color=BLUE, bold=True, align=PP_ALIGN.CENTER)


def make_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    background = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    fill(background, BG)

    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.18))
    fill(header, NAVY)
    text_box(slide, 0.38, 0.14, 12.55, 0.2, "TATA ELXSI  |  TELIPORT SEASON 4  |  CASE STUDY 1", size=8.2, color=CYAN, bold=True)
    text_box(slide, 0.38, 0.36, 12.55, 0.34, "AEROLINK-PHM  |  Closed-Loop Predictive Maintenance for a Connected Fleet", size=18, color=WHITE, bold=True)
    text_box(slide, 0.38, 0.76, 12.55, 0.22, "NASA C-MAPSS • Continuous anomaly + RUL inference • Deterministic non-LLM work-order triage • Human review only for exceptions", size=8.2, color=RGBColor(203, 213, 225))

    text_box(slide, 0.4, 1.29, 12.53, 0.2, "END-TO-END CLOSED LOOP  |  batch benchmark now; streaming adapter ready for fleet telemetry", size=8, color=BLUE, bold=True, align=PP_ALIGN.CENTER)
    nodes = [
        (0.43, "1  INGEST", "Unit, cycle, 3 settings\n+& 21 sensor channels", BLUE),
        (2.92, "2  PREPARE", "Schema / missing-data checks\n+regime-aware features", TEAL),
        (5.41, "3  PREDICT", "Anomaly score + continuous\nRUL + confidence", BLUE),
        (7.90, "4  TRIAGE", "Tunable rules, dedupe\nand exception gates", ORANGE),
        (10.39, "5  ACT & LEARN", "CMMS work order + closure\nfeedback to monitoring", TEAL),
    ]
    for x, title, detail, accent in nodes:
        flow_node(slide, x, title, detail, accent)
    for x in (2.62, 5.11, 7.60, 10.09):
        arrow(slide, x)

    card(slide, 0.40, 2.95, 3.08, 2.96, "DATA ENGINEERING  |  reproducible & regime-aware", TEAL)
    body(slide, 0.55, 3.40, 2.78, [
        ("Validate → ", "schema, duplicate cycles, bounds and missingness; retain a data-quality flag."),
        ("Clean → ", "within-engine forward fill, then training-set median; never leak test/future cycles."),
        ("Normalize → ", "fit scaler by operating-condition cluster for FD002 / FD004; global scaler for single-condition sets."),
        ("Features → ", "rolling mean, standard deviation, slope and deltas (15/30 cycles), computed per engine."),
        ("Label → ", "training RUL = final failure cycle − current cycle; tune any RUL cap on validation only."),
    ])

    card(slide, 3.63, 2.95, 3.08, 2.96, "ML PROGNOSTICS  |  early warning + RUL", BLUE)
    body(slide, 3.78, 3.40, 2.78, [
        ("Anomaly → ", "Isolation Forest calibrated on healthy early-life windows; threshold set to an alert-rate target."),
        ("RUL → ", "gradient-boosted regressor on rolling features; predicts every cycle, not only after an alert."),
        ("Confidence → ", "prediction interval / ensemble spread travels with each RUL estimate."),
        ("Validation → ", "engine-grouped, time-respecting validation; report RMSE + NASA asymmetric score."),
        ("Explain → ", "feature contribution and trend evidence accompany the model output."),
    ])

    card(slide, 6.86, 2.95, 3.08, 2.96, "TRIAGE AGENT  |  deterministic, no LLM", ORANGE)
    body(slide, 7.01, 3.40, 2.78, [
        ("Critical → ", "RUL ≤ 15 cycles AND confidence acceptable: create P1 work order; hold asset per policy."),
        ("Warning → ", "RUL 16–45: create P2 planned-maintenance work order with due-cycle window."),
        ("Watch → ", "RUL 46–80 or anomaly-only: intensify monitoring; no duplicate work order."),
        ("Exception → ", "low confidence, poor data quality or conflicting evidence routes to a reliability engineer."),
        ("Control → ", "state machine uses hysteresis, audit log and idempotent work-order key."),
    ])

    card(slide, 10.09, 2.95, 2.84, 2.96, "MLOPS & SCALE  |  monitor, retrain, govern", TEAL)
    body(slide, 10.24, 3.40, 2.54, [
        ("Track → ", "MLflow records data version, features, parameters, metrics and model approval."),
        ("Monitor → ", "missingness, feature drift, alert rate, RUL error after closure and work-order outcomes."),
        ("Retrain → ", "candidate retrain on drift / labelled failures; promote only if holdout gate improves."),
        ("Scale → ", "start FD001; extend the same contract to FD002–FD004 (1/6 conditions, 1/2 fault modes)."),
    ], size=6.9)

    policy = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.40), Inches(6.15), Inches(12.53), Inches(0.88))
    fill(policy, NAVY_2)
    text_box(slide, 0.60, 6.31, 1.13, 0.2, "WORK-ORDER OUTPUT", size=7.3, color=CYAN, bold=True)
    text_box(slide, 0.60, 6.55, 1.13, 0.22, "automatic & auditable", size=6.3, color=RGBColor(203, 213, 225))
    text_box(slide, 1.88, 6.29, 8.30, 0.43, "{work_order_id, asset_id, priority, predicted_RUL, confidence, anomaly_score, trend_evidence, recommended_due_cycle, model_version}", size=7.4, color=WHITE, font="Aptos Mono")
    text_box(slide, 10.30, 6.27, 2.35, 0.50, "Thresholds are policy parameters: calibrate with maintenance capacity, safety rules and validation results — not assumed ROI.", size=6.2, color=RGBColor(203, 213, 225), align=PP_ALIGN.CENTER)

    footer = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(7.23), prs.slide_width, Inches(0.27))
    fill(footer, NAVY)
    text_box(slide, 0.4, 7.27, 12.5, 0.13, "ROUND 1 COVERAGE: Architecture  •  Data cleaning & rolling features  •  Anomaly + RUL models  •  Rule-based agent  •  MLOps retraining & monitoring", size=6.5, color=RGBColor(203, 213, 225), align=PP_ALIGN.CENTER)

    prs.save(OUTPUT)
    print(f"Created {OUTPUT.resolve()}")


if __name__ == "__main__":
    make_deck()
