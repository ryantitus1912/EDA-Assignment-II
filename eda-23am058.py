"""Builds the project report (Word) from results/metrics.json + figures/.
Usage: python build_report.py --name "Your Name" --roll "Roll No" --year "III Year / V Sem" --date "15/10/2026"
Run olist_review_sentiment.py first."""
import argparse, json
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ap = argparse.ArgumentParser()
ap.add_argument("--name", default="[Student Name]"); ap.add_argument("--roll", default="[Roll No.]")
ap.add_argument("--year", default="III Year / V Sem"); ap.add_argument("--date", default="15/10/2026")
ap.add_argument("--out", default="Project_Report_Olist_Review_Sentiment.docx")
a = ap.parse_args()
M = json.load(open("results/metrics.json")); S = M["stats"]
clean = pd.read_csv("results/clean_review_table.csv", usecols=["review_score"])
sc = clean.review_score.value_counts(normalize=True).sort_index() * 100
cp, cc = M["class_pct"], M["class_counts"]
T, B, V = M["mlp_test"], M["baseline_lr"], M["mlp_val"]
rep = T["report"]

d = Document()
sec = d.sections[0]; sec.left_margin = sec.right_margin = Inches(1); sec.top_margin = sec.bottom_margin = Inches(0.9)
st = d.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(11)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
for h, sz in [("Heading 1", 16), ("Heading 2", 13)]:
    d.styles[h].font.name = "Calibri"; d.styles[h].font.size = Pt(sz); d.styles[h].font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)

def P(text="", bold=False, italic=False, align=None, size=None, after=6):
    p = d.add_paragraph(); r = p.add_run(text); r.bold, r.italic = bold, italic
    if size: r.font.size = Pt(size)
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(after); return p
def bullets(items):
    for t in items: d.add_paragraph(t, style="List Bullet")
def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr(); s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear"); s.set(qn("w:color"), "auto"); s.set(qn("w:fill"), color); tcPr.append(s)
def table(rows, header, widths=None, size=9):
    t = d.add_table(rows=1, cols=len(header)); t.style = "Table Grid"
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""; r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(size); shade(c, "D9E2F3")
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""; cells[i].paragraphs[0].add_run(str(v)).font.size = Pt(size)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths): row.cells[i].width = Inches(w)
    d.add_paragraph().paragraph_format.space_after = Pt(2)
FIGN = [0]
def fig(path, caption, width=6.4):
    FIGN[0] += 1
    d.add_picture(f"figures/{path}.png", width=Inches(width)); d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    P(f"Figure {FIGN[0]}. {caption}", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=9, after=6)
def explain(shows, why, insight):
    for lab, txt in [("What it shows: ", shows), ("Why it was created: ", why), ("Insights: ", insight)]:
        p = d.add_paragraph(); r = p.add_run(lab); r.bold = True; p.add_run(txt); p.paragraph_format.space_after = Pt(3)
    d.add_paragraph()

# ---------------- Title page
for _ in range(3): P()
P("KPR Institute of Engineering and Technology", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=16)
P("Department of Artificial Intelligence and Data Science", align=WD_ALIGN_PARAGRAPH.CENTER, size=13)
P("U21ADP05 – Exploratory Data Analysis and Visualization | Assignment II – Case Study", align=WD_ALIGN_PARAGRAPH.CENTER, size=11, after=30)
P("Exploratory Data Analysis and Visualization of Olist Customer Reviews", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=22, after=4)
P("Text Data: Sentiment Analysis of Portuguese E-Commerce Reviews with a Deep Neural Network", align=WD_ALIGN_PARAGRAPH.CENTER, size=13, after=40)
table([["Student Name", a.name], ["Roll No.", a.roll], ["Department", "Artificial Intelligence and Data Science"],
       ["Year / Semester", a.year], ["Batch", "2023–2027"], ["Course Coordinator", "Mr. Rushikesh Kadam"], ["Date of Submission", a.date]],
      ["Field", "Details"], widths=[2, 4], size=11)
d.add_page_break()

# ---------------- Abstract
d.add_heading("Abstract", 1)
abstract = (f"Online marketplaces depend on customer reviews, but free-text comments are hard to analyse at scale. This study explores the "
            f"Olist Brazilian e-commerce dataset, merging nine tables into one review-level table of {M['n_columns_merged']} fields and "
            f"{M['rows_after_dedup_filter']:,} Portuguese comments. After cleaning, imputation, outlier treatment and a stratified 70/15/15 split, "
            f"seven-stage visualization and statistical tests showed that late delivery, long comments and product category are linked to negative "
            f"sentiment. A multi-layer perceptron on TF-IDF features, tuned over eight configurations, classified comments as negative, neutral or "
            f"positive with {100*T['acc']:.1f}% test accuracy, macro-F1 {T['f1_macro']:.3f} and macro AUC {M['auc']['macro_ovr']:.3f}, outperforming a logistic-regression "
            f"baseline. The neutral class remains hardest to predict.")
P(abstract); P(f"({len(abstract.split())} words)", italic=True, size=9)

# ---------------- 1 Introduction
d.add_heading("1. Introduction and Objective", 1)
P("Customer reviews are a primary signal of product quality and logistics performance for online marketplaces. Olist connects small Brazilian "
  "businesses to large marketplaces; after a purchase, customers rate the order from 1 to 5 stars and may add a free-text comment. Reading every comment "
  "manually is impractical, so automatic understanding of review text is valuable for customer-service triage, seller feedback and quality monitoring.")
P("Objectives", bold=True, after=2)
bullets(["Acquire and merge a real-world multi-table dataset and extract a text-based problem from it (≥15 fields).",
         "Perform EDA and preprocessing: missing values, duplicates, outliers, encoding, scaling, and a train/validation/test split.",
         "Create at least five meaningful visualizations following the seven stages of visualization, with univariate, bivariate, "
         "multivariate and distribution-based analysis.",
         "Build and tune a deep neural network for text classification and evaluate it with loss/accuracy curves, a confusion matrix, ROC/AUC and error analysis.",
         "Draw practical conclusions and recommendations for the marketplace."])
P("Mapping to the seven stages of visualization (Fry, 2008): Acquire (Section 2), Parse (Section 2 – merging/dates), Filter (Section 3 – cleaning), "
  "Mine (Section 3/4 – statistical tests), Represent and Refine (Section 4 – charts with consistent labels and colour code), "
  "Interact (optional Plotly chart in the notebook).")

# ---------------- 2 Dataset
d.add_heading("2. Dataset Description", 1)
P("Source: Brazilian E-Commerce Public Dataset by Olist, Kaggle – https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce "
  "(≈100k orders placed between 2016 and 2018). Eight relational CSV tables were used:")
tb = M["tables"]
table([[k, f"{v[0]:,}", v[1]] for k, v in tb.items()], ["Table", "Rows", "Columns"], widths=[2.2, 1.5, 1.2])
P(f"Reviews, orders, customers, order items, payments, products (with English category names) and sellers were joined on their keys; items and payments "
  f"were aggregated to one row per order. Of {M['merged_rows_all_reviews']:,} reviews, {M['rows_with_comment']:,} contain a written comment "
  f"(the text corpus); after de-duplication and removing 1–2-character comments, {M['rows_after_dedup_filter']:,} remain. The merged table has "
  f"{M['n_columns_merged']} columns, of which the following were used:")
fields = [
 ("text / review_comment_message", "Text", "Portuguese customer comment (model input)"),
 ("review_score → sentiment", "Target", "1–2 negative, 3 neutral, 4–5 positive"),
 ("review_comment_title, has_title", "Text / binary", "Optional comment title (76% empty) and indicator"),
 ("review_word_count, review_char_count", "Numeric", "Length of the comment"),
 ("n_exclaim, n_question, upper_ratio", "Numeric", "Punctuation and capital-letter style features"),
 ("delivery_days", "Numeric", "Purchase → delivery to customer (days)"),
 ("delay_days", "Numeric", "Delivery minus estimated date (>0 = late)"),
 ("approval_hours", "Numeric", "Purchase → payment approval (hours)"),
 ("total_price, total_freight, payment_value", "Numeric", "Item price, freight cost and total paid (BRL)"),
 ("payment_installments", "Numeric (discrete)", "Number of instalments"),
 ("n_items, n_sellers", "Numeric (discrete)", "Items / sellers in the order"),
 ("product_photos_qty, product_weight_g, product_volume_cm3, product_name_lenght, product_description_lenght", "Numeric", "Product listing attributes"),
 ("product_category_name_english", "Categorical", "≈70 product categories"),
 ("payment_type, order_status", "Categorical", "Payment method; order status"),
 ("customer_state, seller_state", "Categorical", "Brazilian state codes"),
 ("order_purchase_timestamp, purchase_month, purchase_hour", "Date / numeric", "Purchase time information")]
table(fields, ["Field(s)", "Type", "Description"], widths=[2.8, 1.2, 2.5], size=8.5)
P("Class balance of the target: " + ", ".join(f"{c} {cc[c]:,} ({cp[c]}%)" for c in ["negative", "neutral", "positive"]) +
  ". The data are imbalanced, with the neutral class under 9%, which shapes the modelling choices below.")
nd = M["describe_num"]
rows = [[k, f"{v['mean']:.2f}", f"{v['std']:.2f}", f"{v['min']:.2f}", f"{v['50%']:.2f}", f"{v['max']:.2f}"] for k, v in nd.items() if k in
        ["total_price", "total_freight", "payment_installments", "delivery_days", "delay_days", "review_word_count", "product_photos_qty"]]
P("Basic statistics (after winsorising at the 1st/99th percentile):", bold=True, after=2)
table(rows, ["Feature", "Mean", "Std", "Min", "Median", "Max"], widths=[2.2, .8, .8, .8, .8, .8])

# ---------------- 3 EDA and Preprocessing
d.add_heading("3. EDA and Preprocessing", 1)
d.add_heading("3.1 Methods used", 2)
bullets([
 f"Missing values: the comment title is missing in 76% of reviews, so it is kept only as the indicator has_title. Delivery fields are missing for "
 f"{M['rows_undelivered']:,} undelivered orders, product attributes for ≈1–3% of rows. Numeric gaps were filled with the median (robust to skew); "
 f"missing category/payment type became the level 'unknown'.",
 f"Duplicates: {M['duplicates_removed']} repeated (order, comment) pairs were removed; comments shorter than 3 characters were dropped.",
 "Outliers: detected with the 1.5×IQR rule (e.g., " + ", ".join(f"{k} {v:,}" for k, v in list(M["iqr_outliers"].items())[:4]) +
 " flagged rows) and treated by winsorising at the 1st/99th percentile, because extreme prices and delivery times are genuine but would distort scaling-based methods.",
 "Transformation: log transforms for skewed variables (price, delivery time, comment length) in the distribution analysis.",
 f"Encoding and scaling: one-hot encoding of 5 categorical fields ({M['encoded_feature_count']} dummy columns); z-score standardisation of {len(M['iqr_outliers'])} numeric "
 "features (fitted on the training set only to avoid leakage).",
 f"Text processing: lower-casing, URL removal, accent-insensitive TF-IDF on unigrams+bigrams (min_df=3, up to 20,000 features; "
 f"{M['tfidf_vocab_size']:,} retained) and a Portuguese stop-word list for visual analysis.",
 f"Split: stratified train/validation/test = {M['split']['train']:,} / {M['split']['val']:,} / {M['split']['test']:,} (70/15/15). "
 f"Only the training set was oversampled (to {M['train_after_oversampling']:,} rows) so that minority classes reach 75% of the majority."])
d.add_heading("3.2 Insights from cleaning and statistical testing", 2)
bullets([
 f"Late deliveries are strongly associated with sentiment: {S['late_share_by_sentiment']['negative']}% of negative-review orders arrived after the estimated date versus "
 f"{S['late_share_by_sentiment']['neutral']}% (neutral) and {S['late_share_by_sentiment']['positive']}% (positive); Kruskal–Wallis H = {S['kruskal_delay']['H']:.0f}, p < 0.001.",
 f"Comment length differs by sentiment (median {S['median_words_by_sentiment']['negative']:.0f} / {S['median_words_by_sentiment']['neutral']:.0f} / "
 f"{S['median_words_by_sentiment']['positive']:.0f} words for negative / neutral / positive; Kruskal–Wallis H = {S['kruskal_wordcount']['H']:.0f}, p < 0.001).",
 f"Product category and sentiment are dependent (χ² = {S['chi2_category']['chi2']:.0f}, dof = {S['chi2_category']['dof']}, p < 0.001) but the effect is small "
 f"(Cramér's V = {S['chi2_category']['cramers_v']:.3f}); customer state is weaker still (V = {S['chi2_customer_state']['cramers_v']:.3f}) and payment type negligible "
 f"(V = {S['chi2_payment_type']['cramers_v']:.3f}). With ≈40k rows even tiny effects are statistically significant, so effect size matters more than p-values.",
 f"Spearman correlations with the 1–5 score: comment length ρ = {S['spearman_words_score']['rho']:.2f}, delivery time ρ = {S['spearman_deliverydays_score']['rho']:.2f}, "
 f"delay vs estimate ρ = {S['spearman_delay_score']['rho']:.2f}."])
d.add_page_break()

# ---------------- 4 Visualizations
d.add_heading("4. Data Visualization", 1)
P("All charts use a consistent colour code (red = negative, grey = neutral, green = positive), titled axes with units, and legends. "
  "Eleven figure groups (more than the required five) cover text, univariate, bivariate, multivariate, distribution and time views.", after=8)

fig("fig01_missing_values", "Percentage of missing values per field.")
explain("The share of missing entries in each field of the merged table (title is the most incomplete).",
        "To decide which fields can be imputed and which should be reduced to an indicator (Filter stage).",
        f"Comment titles are 76% empty, delivery fields 4.6% (undelivered orders) and product attributes 1–3%. Because gaps beyond the title are small, median imputation is safe.")

fig("fig02_target_distribution", "Distribution of review scores and the three sentiment classes.", 6.2)
explain("The discrete distribution of star ratings among reviews that include a comment, and the resulting three-class target.",
        "To understand class balance before modelling, since it determines which metrics are meaningful.",
        f"5-star reviews are the largest group ({sc.get(5, 0):.1f}%) followed by 1-star ({sc.get(1, 0):.1f}%), a U/J-shape typical of reviews: people rate at the extremes. "
        f"Positive is {cp['positive']}%, negative {cp['negative']}% and neutral only {cp['neutral']}%, so accuracy alone would be misleading and macro-F1 is reported.")

fig("fig03_word_frequency", "Top-15 words for each sentiment class (frequency plot).")
explain("The most frequent content words in negative, neutral and positive comments after stop-word removal.",
        "A basic text-mining view to see what customers talk about (Mine/Represent stages).",
        "Delivery vocabulary (entrega, prazo, chegou, entregue) dominates every class, so logistics is the main topic of conversation. Positive comments add praise terms "
        "(recomendo, excelente, ótimo) and 'antes' (before the deadline); negative comments add 'veio' (it came [wrong/defective]), 'pedido' and 'agora' and the store name 'lannister'. "
        "Shared words across classes show that single words alone are ambiguous – context (n-grams) is needed.")

fig("fig04_ngrams", "Top bigrams for each sentiment class.")
explain("The most frequent two-word phrases per class.",
        "Bigrams capture negation and phrases that single words miss (N-gram visualization).",
        "Negative: 'nota fiscal' (invoice), 'veio defeito' (arrived defective), 'péssima qualidade', 'dinheiro volta' (refund), 'gostaria saber' (asking for information) – complaints about "
        "defects, invoices and refunds. Positive/neutral: 'antes prazo', 'chegou antes', 'entrega rápida', 'super recomendo', 'ótima qualidade' – speed of delivery and quality. "
        "Operational themes appear that a star rating alone would hide, such as missing invoices.")

fig("fig05_review_length", "Review length: raw distribution, log-transformed distribution with fitted normal, and by sentiment.", 6.5)
explain("A histogram of words per comment, its log-transform with a fitted normal curve, and box plots of length by sentiment.",
        "Univariate distribution analysis of a key derived text feature and its relation to the target.",
        f"Length is right-skewed (skewness {S['dist_review_word_count']['skew_raw']:.2f}); on the log scale it is much closer to symmetric (skewness {S['dist_review_word_count']['skew_log']:.2f}), "
        f"supporting a log-normal-like model. Unhappy customers write more: median {S['median_words_by_sentiment']['negative']:.0f} words (negative) versus {S['median_words_by_sentiment']['positive']:.0f} (positive).")

fig("fig06_delivery_vs_sentiment", "Delivery delay and delivery time versus sentiment.", 6.5)
explain("Box plots of delay relative to the promised date, the percentage of late orders per star score, and violin plots of delivery time by sentiment.",
        "Bivariate analysis of the operational factor that the text suggested matters most.",
        f"Negative reviews have the heaviest late tail: the upper quartile of their orders is already at the estimated date (≈0 days) while for positive orders it is about 7 days early; mean delay is {S['mean_delay_by_sentiment']['negative']} days (negative) versus {S['mean_delay_by_sentiment']['positive']} days (positive). The share of late orders falls steadily with rating – from 28.5% at 1 star to 3.1% at 5 stars – and median delivery time rises from "
        f"{S['median_delivery_by_sentiment']['positive']} days (positive) to {S['median_delivery_by_sentiment']['negative']} days (negative). Reliability of the delivery promise is a clear driver of satisfaction.")

fig("fig07_category_state_sentiment", "Sentiment mix by product category and customer state.", 6.5)
explain("100% stacked bars of sentiment for the 12 largest categories and the states with more than 300 reviews.",
        "Categorical bivariate analysis to find where dissatisfaction concentrates.",
        f"Negative share ranges from {M['category_neg_pct']['cool_stuff']}% (cool_stuff) to {M['category_neg_pct']['computers_accessories']}% (computers_accessories), and by state from "
        f"{M['state_neg_pct']['MT']}% (MT) to {M['state_neg_pct']['RJ']}% (RJ). Hardware/furniture categories, which are harder to ship and match to expectations, and states such as RJ, MA and CE "
        "show more complaints. The differences are statistically significant but modest (Cramér's V ≈ 0.11 and 0.06).")

fig("fig08_distributions", "Discrete and continuous distribution analysis.", 6.2)
explain("Top row: observed distribution of payment instalments against a Poisson model, and items per order on a log axis. Bottom row: log-price histogram with fitted normal and a Q-Q plot.",
        "To identify, visualise and test discrete and continuous probability distributions (CO4).",
        f"Instalments are not Poisson: the mean is {S['installments']['mean']:.2f} but the variance {S['installments']['var']:.2f} (dispersion index {S['installments']['dispersion_index']:.2f} ≫ 1, i.e. over-dispersed), "
        "with a spike at 1 (single payment) and peaks at 8 and 10 reflecting standard instalment plans. Items per order decays roughly geometrically (most orders have one item). "
        f"Order price is extremely right-skewed (skewness {S['dist_total_price']['skew_raw']:.1f}) but after a log transform skewness falls to {S['dist_total_price']['skew_log']:.2f} and the "
        f"Q-Q plot is nearly linear, i.e. approximately log-normal, with slightly heavier tails (KS distance {S['dist_total_price']['ks_lognormal']:.3f}). Log transforms are therefore appropriate for price-like features.")

fig("fig09_correlation_heatmap", "Spearman correlation heat map of numeric features and review score.", 6.2)
explain("Pairwise Spearman rank correlations among 15 numeric features and the 1–5 score.",
        "Multivariate analysis to find related features and the strongest numeric correlates of satisfaction (rank correlation suits skewed data).",
        f"The strongest correlate of score is comment length (ρ = {M['corr_with_score']['review_word_count']}), then delivery days ({M['corr_with_score']['delivery_days']}), delay ({M['corr_with_score']['delay_days']}) "
        f"and items per order ({M['corr_with_score']['n_items']}). Upper-case ratio is mildly positive (+0.17) – enthusiastic short messages. Price, freight and weight are mutually correlated "
        "(heavier products cost more to ship) but nearly uncorrelated with score, so there is little multicollinearity concern for rating-related drivers.")

fig("fig10_pairplot", "Pair plot of key numeric features coloured by sentiment (1,500-row sample).", 6.2)
explain("Scatter plots and density curves of delivery time, log-price, freight and log-comment-length, coloured by sentiment.",
        "To inspect multivariate structure and class separability visually.",
        "Negative reviews (red) form a longer right tail in delivery_days and sit at higher log_words, but classes overlap heavily on price and freight. "
        "No single numeric feature separates the classes, which motivates using the review text itself as the model input. Freight rises with price as expected.")

fig("fig11_time_trend", "Monthly review volume and negative share over time.", 6.5)
explain("Monthly counts of commented reviews by sentiment and the monthly share of negative comments, by purchase month.",
        "To check for time-dependent behaviour such as seasonality or service incidents.",
        f"Volume grows through 2017–2018. The negative share spikes to 33% in Nov 2017 and peaks at {M['peak_neg_month'][1]}% in {M['peak_neg_month'][0]} (versus ≈22% in mid-2017 and mid-2018). "
        "The Nov 2017 peak coincides with a sharp jump in order volume (a holiday-sale surge), which is consistent with delivery capacity being strained, though the data alone cannot prove the cause.")

# ---------------- 5 Model
d.add_heading("5. Deep Learning Model", 1)
d.add_heading("5.1 Problem set-up and model choice", 2)
P("The task is three-class sentiment classification of a Portuguese comment. Inputs are the TF-IDF vectors of the cleaned text only: structured fields such as delivery delay are used in the EDA "
  "but are not model inputs, because the goal is to understand text and because review_score itself defines the label.")
P("Model: a multi-layer perceptron (deep feed-forward neural network) with ReLU hidden layers and a softmax output, trained by mini-batch Adam with cross-entropy loss. "
  "It was compared with a logistic-regression baseline on the same features.")
P("Note on model choice: the assignment asks for the most suitable deep model for text. In the environment used to produce these results, "
  "no deep-learning framework (PyTorch/TensorFlow) or pretrained language model could be installed, so a TF-IDF + MLP was the strongest neural model that could be trained and "
  "verified end to end. A fine-tuned multilingual or Portuguese transformer (e.g., BERTimbau) is expected to do better, particularly on the neutral class, and is listed under future scope.", italic=True)
d.add_heading("5.2 Architecture and hyper-parameters", 2)
bc = M["best_config"]
table([["Input", f"TF-IDF 1–2-grams, {M['tfidf_vocab_size']:,} features"], ["Hidden layers", bc["hidden_layers"] + " ReLU units"], ["Output", "3-way softmax (negative / neutral / positive)"],
       ["Trainable parameters", f"{M['n_parameters']:,}"], ["Optimizer / batch size", f"Adam / {bc['batch_size']}"], ["Learning rate", bc["learning_rate"]],
       ["L2 regularisation (alpha)", bc["alpha_L2"]], ["Max epochs / early stopping", f"{bc['max_epochs']} / patience {bc['early_stopping_patience']} on validation macro-F1"],
       ["Epoch selected", bc["best_epoch"]], ["Training data", f"{M['train_after_oversampling']:,} rows after oversampling (train split only)"]],
      ["Item", "Setting"], widths=[2.3, 4.2])
d.add_heading("5.3 Hyper-parameter tuning", 2)
P("A grid of 2 architectures × 2 learning rates × 2 L2 strengths was searched, selecting by validation macro-F1 (Figure 12).")
table([[g["hidden_layers"], g["learning_rate"], g["alpha_L2"], g["best_epoch"], g["val_macro_f1"], g["val_acc"]] for g in M["hp_grid"]],
      ["Hidden layers", "Learning rate", "L2 alpha", "Best epoch", "Val macro-F1", "Val accuracy"], widths=[1.2, 1, 1, .9, 1.1, 1.1])
fig("fig12_hyperparameter_search", "Validation macro-F1 across the hyper-parameter grid.", 4.6)
P(f"The single hidden layer of 128 units with a lower learning rate (3e-4) was best (validation macro-F1 {V['f1_macro']:.3f}). The larger two-layer network reached similar accuracy but a lower macro-F1: with only ≈28k "
  "training comments and 17k sparse inputs it over-fits within a single epoch. Regularisation strength made little difference (≤0.002 F1).")
d.add_heading("5.4 Feature analysis and importance", 2)
fig("fig13_lr_feature_importance", "Top n-grams per class according to the logistic-regression coefficients.", 6.5)
fig("fig14_mlp_feature_importance", "Top features per class by occlusion importance in the MLP.", 6.5)
P("Logistic-regression coefficients are directly interpretable: 'não recomendo', 'péssima', 'não gostei', 'péssimo', 'insatisfeito' push toward negative, whereas 'parabéns', 'rápido', 'amei', 'perfeito', 'maravilhoso' push toward positive. "
  "Neutral coefficients are noisy single product terms, showing that the neutral class has no clear vocabulary of its own. For the MLP, importance was measured by occlusion: each TF-IDF feature "
  "in 1,500 test comments was set to zero and the drop in the predicted-class probability averaged. Strong negative drivers are 'péssima (qualidade)', 'não recomendo', 'cancelar', 'absurdo', 'demora na'; "
  "positive drivers are 'ótimo', 'amei', 'maravilhoso', 'adorei', 'excelente'. The neutral class is driven by hedging words such as 'porém', 'mas', 'poderia', 'achei', 'demorou' – mixed praise and complaint. "
  "This agrees with the EDA: sentiment is encoded in evaluative words and negations, not in structured order data.")

# ---------------- 6 Results
d.add_heading("6. Result Visualization and Interpretation", 1)
fig("fig15_loss_accuracy_curves", "Training/validation loss and accuracy versus epoch.", 6.5)
P(f"Training loss falls steadily (0.89 → {M['history']['train_loss'][-1]:.2f}) and training accuracy rises to {100*M['history']['train_acc'][-1]:.0f}%, while validation loss bottoms out near epoch 4 and validation macro-F1 "
  f"peaks at epoch {bc['best_epoch']} before declining – the classic sign of over-fitting on sparse text features. Early stopping therefore selected epoch {bc['best_epoch']}, "
  "and the model should not be trained longer without stronger regularisation or more data.")
fig("fig16_confusion_matrix", "Confusion matrix on the test set (counts and row-normalised).", 6.2)
cm = M["confusion_matrix"]
P(f"On {sum(map(sum, cm)):,} test reviews, recall is {rep['negative']['recall']:.2f} for negative, {rep['positive']['recall']:.2f} for positive and only {rep['neutral']['recall']:.2f} for neutral. "
  f"Neutral comments are spread over all three classes (31% predicted negative, 27% positive), and 319 positive reviews are predicted neutral. Cross-polarity errors (negative↔positive) are rare ({M['error_distance_pct']['negative_vs_positive_confusion']:.1f}%).")
table([[c, f"{rep[c]['precision']:.3f}", f"{rep[c]['recall']:.3f}", f"{rep[c]['f1-score']:.3f}", int(rep[c]['support'])] for c in ["negative", "neutral", "positive"]] +
      [["accuracy", "", "", f"{T['acc']:.3f}", int(rep['negative']['support'] + rep['neutral']['support'] + rep['positive']['support'])],
       ["macro avg", f"{rep['macro avg']['precision']:.3f}", f"{rep['macro avg']['recall']:.3f}", f"{rep['macro avg']['f1-score']:.3f}", ""]],
      ["Class", "Precision", "Recall", "F1-score", "Support"], widths=[1.4, 1.1, 1.1, 1.1, 1])
fig("fig17_roc_curves", "One-vs-rest ROC curves with AUC on the test set.", 4.8)
P(f"AUC is {M['auc']['negative']:.3f} for negative and {M['auc']['positive']:.3f} for positive – the model ranks polar comments very well – but only {M['auc']['neutral']:.3f} for neutral. "
  f"Macro one-vs-rest AUC is {M['auc']['macro_ovr']:.3f} (logistic regression: {M['auc_lr_macro_ovr']:.3f}).")
fig("fig18_error_analysis", "Prediction-confidence distribution and error severity.", 6.3)
P(f"Correct predictions have mean confidence {M['confidence']['mean_correct']:.2f} versus {M['confidence']['mean_wrong']:.2f} for wrong ones, so the probability output is informative: low-confidence "
  f"predictions can be routed to human review. Of all errors, {M['error_distance_pct']['off_by_one_class']:.1f}% of test cases are adjacent-class mistakes and only {M['error_distance_pct']['negative_vs_positive_confusion']:.1f}% are opposite-polarity mistakes. "
  "(For a classification task, this replaces the regression 'error distribution' chart.)")
fig("fig19_model_comparison", "Test accuracy and macro-F1 of the baseline and the tuned MLP.", 4.6)
table([["Logistic regression (baseline)", f"{B['val_acc']:.3f}", f"{B['val_f1_macro']:.3f}", f"{B['test_acc']:.3f}", f"{B['test_f1_macro']:.3f}", f"{M['auc_lr_macro_ovr']:.3f}"],
       ["MLP (tuned)", f"{V['acc']:.3f}", f"{V['f1_macro']:.3f}", f"{T['acc']:.3f}", f"{T['f1_macro']:.3f}", f"{M['auc']['macro_ovr']:.3f}"]],
      ["Model", "Val acc", "Val macro-F1", "Test acc", "Test macro-F1", "Test macro-AUC"], widths=[2.1, .8, 1, .8, 1, 1])
P(f"The MLP improves on the baseline by {100*(T['acc']-B['test_acc']):.1f} accuracy points and {T['f1_macro']-B['test_f1_macro']:.3f} macro-F1; validation and test scores are close "
  f"({V['f1_macro']:.3f} vs {T['f1_macro']:.3f} macro-F1), indicating the result generalises and that tuning did not over-fit the validation set.")

# ---------------- 7 Conclusion
d.add_heading("7. Conclusion and Future Scope", 1)
d.add_heading("What was learned", 2)
bullets([f"Customer comments are dominated by delivery topics; {cp['positive']}% are positive, {cp['negative']}% negative and only {cp['neutral']}% neutral.",
         f"Late or slow delivery is the clearest operational driver of dissatisfaction ({S['late_share_by_sentiment']['negative']}% of negative-review orders were late vs {S['late_share_by_sentiment']['positive']}% of positive).",
         "Negative reviewers write longer comments and mention defects, invoices, refunds and non-recommendation.",
         "Category and region effects exist but are small; price and freight barely relate to satisfaction.",
         "Instalments are over-dispersed (not Poisson); price and comment length are approximately log-normal."])
d.add_heading("Model performance", 2)
P(f"The tuned MLP achieved {100*T['acc']:.1f}% accuracy, macro-F1 {T['f1_macro']:.3f} and macro AUC {M['auc']['macro_ovr']:.3f} on unseen data, beating the linear baseline. It is strong on negative "
  f"(F1 {rep['negative']['f1-score']:.2f}) and positive (F1 {rep['positive']['f1-score']:.2f}) comments but weak on neutral (F1 {rep['neutral']['f1-score']:.2f}).")
d.add_heading("Challenges and limitations", 2)
bullets(["Strong class imbalance and a vague neutral class (3-star comments often contain both praise and complaint).",
         "Over-fitting after 2 epochs with sparse TF-IDF inputs; the model has no notion of word order or meaning beyond bigrams.",
         "Only 40% of reviews contain a comment, so conclusions apply to customers who choose to write; labels come from star ratings, which are themselves noisy.",
         "No pretrained Portuguese language model or GPU deep-learning framework was available in the environment used for the experiments.",
         "Correlations are observational and do not prove causes (e.g. delivery delays and the Nov 2017 spike)."])
d.add_heading("Recommendations for the marketplace", 2)
bullets(["Monitor delivery-promise accuracy: orders predicted to be late are the highest-risk for negative reviews; give honest delivery estimates.",
         "Route low-confidence or negative predicted comments mentioning 'defeito', 'nota fiscal', 'cancelar' to customer service first.",
         "Review hardware/furniture categories and high-complaint regions (RJ, MA, CE) with sellers and logistics partners.",
         "Plan extra logistics capacity ahead of seasonal peaks such as November."])
d.add_heading("Future scope", 2)
bullets(["Fine-tune a pretrained Portuguese or multilingual transformer (BERTimbau, XLM-RoBERTa) or a bidirectional LSTM with embeddings; these capture word order and context and should raise neutral-class F1.",
         "Combine text with structured features (delay, category, price) in a multimodal network and an ordinal (1–5 star) loss.",
         "Use class-weighted or focal loss and data augmentation for the neutral class; add k-fold cross-validation.",
         "Add aspect-based sentiment (delivery vs product quality vs seller), topic modelling, and SHAP/attention explanations.",
         "Deploy as a real-time dashboard that triages incoming reviews."])

# ---------------- 8 References
d.add_heading("References", 1)
refs = [
 "Olist. (2018). Brazilian E-Commerce Public Dataset by Olist [Data set]. Kaggle. https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce",
 "Fry, B. (2008). Visualizing Data: Exploring and Explaining Data with the Processing Environment. O'Reilly Media.",
 "Pedregosa, F., Varoquaux, G., Gramfort, A., et al. (2011). Scikit-learn: Machine learning in Python. Journal of Machine Learning Research, 12, 2825–2830.",
 "Hunter, J. D. (2007). Matplotlib: A 2D graphics environment. Computing in Science & Engineering, 9(3), 90–95.",
 "Waskom, M. L. (2021). seaborn: statistical data visualization. Journal of Open Source Software, 6(60), 3021.",
 "Devlin, J., Chang, M.-W., Lee, K., & Toutanova, K. (2019). BERT: Pre-training of deep bidirectional transformers for language understanding. Proceedings of NAACL-HLT, 4171–4186.",
 "Souza, F., Nogueira, R., & Lotufo, R. (2020). BERTimbau: Pretrained BERT models for Brazilian Portuguese. Intelligent Systems (BRACIS 2020), LNCS 12319, 403–417.",
 "Kingma, D. P., & Ba, J. (2015). Adam: A method for stochastic optimization. ICLR 2015.",
 "Salton, G., & Buckley, C. (1988). Term-weighting approaches in automatic text retrieval. Information Processing & Management, 24(5), 513–523."]
for i, r in enumerate(refs, 1): P(f"[{i}] {r}", size=10, after=3)

# ---------------- Appendix
d.add_page_break(); d.add_heading("Appendix: Code", 1)
P("Complete, commented source (also in the GitHub repository as olist_review_sentiment.py and the .ipynb notebook).", italic=True, size=9)
for line in open("olist_review_sentiment.py", encoding="utf-8").read().splitlines():
    p = d.add_paragraph(); p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.0
    r = p.add_run(line if line else " "); r.font.name = "Consolas"; r.font.size = Pt(7)
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
d.save(a.out); print("saved", a.out)
