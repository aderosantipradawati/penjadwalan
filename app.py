import calendar
import datetime as dt
import html
import math
import random
from io import BytesIO

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

try:
    import docx  # python-docx
except ImportError:
    docx = None

WEEKDAYS = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
MONTH_NAMES = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]


# ----------------------------------------------------------------------
# Label Hari (helper)
# ----------------------------------------------------------------------
def get_day_keys(scope: str, month: int = None, year: int = None):
    if scope == "Seminggu":
        return list(WEEKDAYS), list(WEEKDAYS)
    n_days = calendar.monthrange(year, month)[1]
    keys, labels = [], []
    for day in range(1, n_days + 1):
        date = dt.date(year, month, day)
        keys.append(date.isoformat())
        hari_singkat = WEEKDAYS[date.weekday()][:3]
        nama_bulan = MONTH_NAMES[date.month - 1]
        labels.append(f"{hari_singkat}, {day} {nama_bulan} {year}")
    return keys, labels


def dates_in_range(start: dt.date, end: dt.date):
    if start > end:
        start, end = end, start
    days = (end - start).days
    return [(start + dt.timedelta(days=i)).isoformat() for i in range(days + 1)]

# ----------------------------------------------------------------------
# Session state
# ----------------------------------------------------------------------
def init_state():
    today = dt.date.today()
    defaults = {
        "item_list": [],
        "banned_days": {},
        "excluded_days": set(),   # hari/tanggal libur — tidak ada jadwal sama sekali
        "exclusion_pairs": [],
        "scope": "Seminggu",
        "month": today.month,
        "year": today.year,
        "capacity_mode": "Otomatis",
        "capacity_manual": 1,
        "capacity_target_mode": "Tepat",  # "Tepat" = wajib pas; "Maksimal" = batas atas, boleh kurang
        "result": None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


init_state()

st.set_page_config(page_title="Penjadwalan", layout="centered")

# Sembunyikan menu bawaan Streamlit (hamburger menu, tombol Deploy, footer, dsb.)
# supaya tampil seperti aplikasi mandiri, terutama di layar HP.
st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden; height: 0;}
    [data-testid="stToolbar"] {visibility: hidden; height: 0; position: fixed;}
    [data-testid="stDecoration"] {display: none;}
    [data-testid="stStatusWidget"] {visibility: hidden;}
    .stDeployButton {display: none;}
    div.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <style>
    [data-testid="stHeadingWithActionElements"] {
        justify-content: center;
        text-align: center;
    }
    </style>
    """,
    unsafe_allow_html=True
)
if st.session_state.scope == "Seminggu":
    st.title("Jadwal Mingguan", help="Susun jadwal secara otomatis untuk satu minggu dengan mempertimbangkan hari yang tidak boleh digunakan, item yang tidak boleh muncul di hari yang sama, dan target jumlah item per hari.")
else:
    st.title("Jadwal Bulanan", help="Susun jadwal secara otomatis untuk satu bulan tertentu, dengan mempertimbangkan tanggal yang tidak boleh digunakan, item yang tidak boleh muncul di tanggal yang sama, dan target jumlah item per hari.")

# ----------------------------------------------------------------------
# Item (helper)
# ----------------------------------------------------------------------
def add_item(name: str):
    name = name.strip()
    if name and name not in st.session_state.item_list:
        st.session_state.item_list.append(name)
        st.session_state.banned_days[name] = set()


def remove_item(name: str):
    if name in st.session_state.item_list:
        st.session_state.item_list.remove(name)
        st.session_state.banned_days.pop(name, None)
        st.session_state.exclusion_pairs = [
            p for p in st.session_state.exclusion_pairs if name not in p
        ]


def dedupe_preserve_order(values):
    seen, out = set(), []
    for v in values:
        v = str(v).strip()
        if v and v.lower() != "nan" and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def parse_excel(file) -> list:
    xls = pd.ExcelFile(file)
    values = []
    for sheet in xls.sheet_names:
        df = xls.parse(sheet, header=None)
        for col in df.columns:
            values.extend(df[col].dropna().tolist())
    return dedupe_preserve_order(values)


def parse_docx(file) -> list:
    if docx is None:
        st.error("python-docx belum dinstal — tambah 'python-docx' ke requirements.txt.")
        return []
    document = docx.Document(file)
    values = [p.text for p in document.paragraphs]
    return dedupe_preserve_order(values)


def handle_upload(file, parse_fn, state_key):
    if file is None:
        return
    marker = f"{file.name}_{file.size}"
    marker_key = f"_last_{state_key}"
    if st.session_state.get(marker_key) == marker:
        return
    names = parse_fn(file)
    added = 0
    for n in names:
        if n not in st.session_state.item_list:
            add_item(n)
            added += 1
    st.session_state[marker_key] = marker
    st.success(f"Ditemukan {len(names)} item unik di file, ditambahkan {added} item baru.")


# ----------------------------------------------------------------------
# Kapasitas & Kemunculan
# ----------------------------------------------------------------------
def compute_capacity(n_items, n_days, mode, manual_value):
    if n_items == 0 or n_days == 0:
        return 0
    if mode == "Manual":
        return max(1, int(manual_value))
    return max(1, math.ceil(n_items / n_days))


def compute_occurrences(items, n_days, capacity, target_mode="Tepat"):
    """Tepat: total slot (kapasitas × hari) dibagi ke item — sisa bagi jatuh ke
    ITEM, jadi tiap hari pas sama kapasitas, tapi jumlah kemunculan antar item
    bisa beda (misal ada yang 2x, ada yang 3x).
    Maksimal: tiap item dapat jumlah kemunculan yang SAMA RATA (adil) — sisa
    bagi jatuh ke HARI nanti (lihat run_ga), jadi total per hari bisa sedikit
    berbeda (misal kadang 5, kadang 6) menyesuaikan pembagian rata itu."""
    n = len(items)
    if n == 0 or n_days == 0:
        return {}
    total_slots = capacity * n_days
    if target_mode == "Maksimal":
        occ_per_item = max(1, round(total_slots / n))
        return {item: occ_per_item for item in items}
    base, rem = divmod(total_slots, n)
    return {item: base + (1 if i < rem else 0) for i, item in enumerate(items)}


def build_tokens(items, occurrences):
    tokens = []
    for item in items:
        tokens.extend([item] * occurrences.get(item, 0))
    return tokens


# ----------------------------------------------------------------------
# GA
# ----------------------------------------------------------------------
def run_ga(tokens, day_keys, banned_days, exclusion_pairs, capacity,
           population_size=150, generations=300, mutation_rate=0.15, target_mode="Tepat"):
    n_tokens = len(tokens)
    n_days = len(day_keys)
    day_pos = {k: i for i, k in enumerate(day_keys)}
    banned_idx = {
        item: {day_pos[d] for d in banned_days.get(item, set()) if d in day_pos}
        for item in set(tokens)
    }
    pair_list = list(exclusion_pairs)

    # Untuk mode Maksimal: karena jumlah kemunculan item sudah dipastikan SAMA
    # RATA (lihat compute_occurrences), sisa pembagian sekarang jatuh ke hari —
    # jadi target realistis per hari adalah rentang [low, high], bukan angka
    # kapasitas mentah. Contoh: 40 token / 7 hari -> low=5, high=6 (5 hari
    # dapat 6, 2 hari dapat 5). Kedua angka itu SAH, tidak ada penalti.
    if n_days:
        day_base, day_rem = divmod(n_tokens, n_days)
    else:
        day_base, day_rem = 0, 0
    day_low = day_base
    day_high = day_base + (1 if day_rem else 0)

    def random_chromo():
        return [random.randrange(n_days) for _ in range(n_tokens)] if n_days else []

    def fitness(chromo):
        score = 1_000_000
        day_items = [[] for _ in range(n_days)]
        for idx, d in enumerate(chromo):
            day_items[d].append(tokens[idx])

        for idx, d in enumerate(chromo):
            if d in banned_idx.get(tokens[idx], set()):
                score -= 1000

        for items_on_day in day_items:
            counts = {}
            for it in items_on_day:
                counts[it] = counts.get(it, 0) + 1
            for c in counts.values():
                if c > 1:
                    score -= 800 * (c - 1)

        day_sets = [set(x) for x in day_items]
        for a, b in pair_list:
            for s in day_sets:
                if a in s and b in s:
                    score -= 1500

        for items_on_day in day_items:
            count = len(items_on_day)
            if target_mode == "Maksimal":
                if count < day_low:
                    score -= (day_low - count) * 300
                elif count > day_high:
                    score -= (count - day_high) * 300
                # day_low <= count <= day_high: dalam rentang wajar, tidak ada penalti
            else:
                score -= abs(count - capacity) * 300

        return score


    def tournament(pop, k=5):
        return max(random.sample(pop, min(k, len(pop))), key=lambda ind: ind["fitness"])

    population = [{"chromo": random_chromo(), "fitness": 0} for _ in range(population_size)]
    history = []

    for gen in range(generations):
        for ind in population:
            ind["fitness"] = fitness(ind["chromo"])

        if gen % 10 == 0:
            best = max(population, key=lambda ind: ind["fitness"])
            history.append({"gen": gen, "score": best["fitness"]})

        population.sort(key=lambda ind: ind["fitness"], reverse=True)
        elite_count = max(1, population_size // 10)
        new_pop = population[:elite_count]

        while len(new_pop) < population_size:
            p1 = tournament(population)["chromo"]
            p2 = tournament(population)["chromo"]
            child = [p1[i] if random.random() < 0.5 else p2[i] for i in range(n_tokens)]
            if n_tokens and random.random() < mutation_rate:
                child[random.randrange(n_tokens)] = random.randrange(n_days)
            new_pop.append({"chromo": child, "fitness": 0})

        population = new_pop

    for ind in population:
        ind["fitness"] = fitness(ind["chromo"])
    best = max(population, key=lambda ind: ind["fitness"])
    history.append({"gen": generations, "score": best["fitness"]})

    return best["chromo"], best["fitness"], history


def diagnose(tokens, chromo, day_keys, day_labels, banned_days, exclusion_pairs, capacity, target_mode="Tepat"):
    violations = []
    day_items = [[] for _ in day_keys]
    for idx, d in enumerate(chromo):
        day_items[d].append(tokens[idx])

    for idx, d in enumerate(chromo):
        item = tokens[idx]
        if day_keys[d] in banned_days.get(item, set()):
            violations.append(f"🚨 **{item}** tidak boleh dijadwalkan pada **{day_labels[d]}**.")

    for i, items_on_day in enumerate(day_items):
        counts = {}
        for it in items_on_day:
            counts[it] = counts.get(it, 0) + 1
        for it, c in counts.items():
            if c > 1:
                violations.append(f"🚨 **{it}** muncul **{c} kali** di **{day_labels[i]}**. Seharusnya hanya 1 kali.")

    day_sets = [set(x) for x in day_items]
    for a, b in exclusion_pairs:
        for i, s in enumerate(day_sets):
            if a in s and b in s:
                violations.append(f"⚠️ **{a}** dan **{b}** dijadwalkan di **{day_labels[i]}**. Keduanya harus berada di hari yang berbeda.")

    if day_keys:
        day_base, day_rem = divmod(len(tokens), len(day_keys))
    else:
        day_base, day_rem = 0, 0
    day_low = day_base
    day_high = day_base + (1 if day_rem else 0)

    for i, items_on_day in enumerate(day_items):
        count = len(items_on_day)
        if target_mode == "Maksimal":
            if count < day_low:
                violations.append(f"⚠️ **{day_labels[i]}** memiliki **{count} item**, di bawah rentang wajar ({day_low}-{day_high} item).")
            elif count > day_high:
                violations.append(f"⚠️ **{day_labels[i]}** memiliki **{count} item**, melebihi rentang wajar ({day_low}-{day_high} item).")
        elif count != capacity:
            violations.append(f"⚠️ **{day_labels[i]}** memiliki **{count} item**. Targetnya **{capacity} item**.")

    if not violations:
        violations.append("✅ Tidak ada masalah dengan jadwal. Semua aturan sudah terpenuhi.")
    return violations


def get_calendar_weeks(scope, year, month, avail_items, excluded_days):
    """Struktur data bersama untuk tampilan kalender di layar MAUPUN pdf,
    supaya keduanya selalu identik. Mengembalikan list minggu; tiap minggu
    adalah list 7 sel (Senin..Minggu), tiap sel dict berisi:
      - header_markup: teks header (pakai <br/> untuk baris baru)
      - content_markup: isi sel (pakai <br/>, dan <font color="..."> untuk warna)
      - is_blank: True kalau sel ini di luar bulan (khusus tampilan bulanan)
    """
    def content_for(key):
        if key in excluded_days:
            return '<font color="#dc3545">(Libur)</font>'
        its = avail_items.get(key, [])
        if its:
            return "<br/>".join(f"&bull; {html.escape(it)}" for it in its)
        return '<font color="#aaaaaa">-</font>'

    if scope == "Seminggu":
        week = []
        for label in WEEKDAYS:
            week.append({
                "header_markup": label,
                "content_markup": content_for(label),
                "is_blank": False,
            })
        return [week]

    cal = calendar.Calendar(firstweekday=0)  # 0 = Senin
    weeks_raw = cal.monthdayscalendar(year, month)
    weeks = []
    for week_nums in weeks_raw:
        week = []
        for i, day_num in enumerate(week_nums):
            if day_num == 0:
                week.append({"header_markup": "", "content_markup": "", "is_blank": True})
                continue
            key = dt.date(year, month, day_num).isoformat()
            week.append({
                "header_markup": f"{WEEKDAYS[i]}<br/>{day_num}",
                "content_markup": content_for(key),
                "is_blank": False,
            })
        weeks.append(week)
    return weeks


def build_calendar_html(weeks):
    """Grid kalender di layar: blok 7 kolom header lalu baris isi, berulang
    per minggu."""
    rows_html = []
    for week in weeks:
        header_cells, content_cells = [], []
        for cell in week:
            if cell["is_blank"]:
                header_cells.append(
                    '<th style="background:#e9ecef;border:1px solid #ddd;padding:6px;min-width:100px;"></th>'
                )
                content_cells.append(
                    '<td style="background:#fafafa;border:1px solid #ddd;padding:6px;"></td>'
                )
                continue
            header_cells.append(
                '<th style="background:#222831;color:#ffbe33;border:1px solid #444;'
                f'padding:6px;min-width:100px;text-align:center;">{cell["header_markup"]}</th>'
            )
            content_cells.append(
                '<td style="border:1px solid #ddd;padding:6px;vertical-align:top;'
                f'font-size:13px;">{cell["content_markup"].replace("&bull;", "•")}</td>'
            )
        rows_html.append("<tr>" + "".join(header_cells) + "</tr>")
        rows_html.append("<tr>" + "".join(content_cells) + "</tr>")

    return (
        '<div style="overflow-x:auto;">'
        '<table style="width:100%; border-collapse:collapse;">'
        + "".join(rows_html) +
        "</table></div>"
    )


def build_calendar_pdf(weeks, scope_label, capacity, target_mode):
    """PDF dengan grid yang sama persis strukturnya dengan tampilan di layar:
    blok 7 kolom header lalu baris isi, berulang per minggu."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(letter), title="Jadwal",
        leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24,
    )
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle("header", parent=styles["Normal"], fontSize=9, leading=11,
                                   textColor=colors.HexColor("#ffbe33"), alignment=1)
    content_style = ParagraphStyle("content", parent=styles["Normal"], fontSize=8, leading=11)

    mode_text = "Tepat" if target_mode == "Tepat" else "Maksimal"
    story = [
        Paragraph("Jadwal", styles["Title"]),
        Paragraph(f"{scope_label} &middot; target kapasitas: {capacity} item/hari ({mode_text})", styles["Normal"]),
        Spacer(1, 10),
    ]

    usable_width = landscape(letter)[0] - 48
    col_width = usable_width / 7

    for week in weeks:
        header_row, content_row = [], []
        for cell in week:
            if cell["is_blank"]:
                header_row.append("")
                content_row.append("")
                continue
            header_row.append(Paragraph(cell["header_markup"], header_style))
            content_row.append(Paragraph(cell["content_markup"], content_style))

        tbl = Table([header_row, content_row], colWidths=[col_width] * 7)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#222831")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 4))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
tab_scope, tab_items, tab_constraints, tab_run = st.tabs(
    ["Periode", "Item", "Aturan", "Jadwalkan"]
)

# --- Tab 1: Periode ---------------------------------------------------------
with tab_scope:
    st.subheader("Pilih periode jadwal")
    scope = st.radio("Periode", ["Seminggu", "Sebulan"], horizontal=True,
                      index=0 if st.session_state.scope == "Seminggu" else 1)
    if scope != st.session_state.scope:
        st.session_state.scope = scope
        st.session_state.result = None
        st.session_state.banned_days = {}
        st.session_state.excluded_days = set()

    if scope == "Sebulan":
        c1, c2 = st.columns(2)
        month = c1.selectbox("Bulan", list(range(1, 13)),
                              format_func=lambda m: MONTH_NAMES[m - 1],
                              index=st.session_state.month - 1)
        year = c2.number_input("Tahun", min_value=2000, max_value=2100, value=st.session_state.year)
        if month != st.session_state.month or year != st.session_state.year:
            st.session_state.month = month
            st.session_state.year = year
            st.session_state.result = None
        n_days = calendar.monthrange(int(year), int(month))[1]
        st.caption(f"{MONTH_NAMES[month - 1]} {year} memiliki {n_days} hari.")
    else:
        st.caption("Menggunakan 7 hari dalam seminggu.")

day_keys, day_labels = get_day_keys(
    st.session_state.scope, st.session_state.month, st.session_state.year
)
label_of = dict(zip(day_keys, day_labels))

avail_day_keys = [k for k in day_keys if k not in st.session_state.excluded_days]
avail_day_labels = [label_of[k] for k in avail_day_keys]

# --- Tab 2: Item ----------------------------------------------------------
with tab_items:
    st.subheader("Tambah item")
    col1, col2 = st.columns([4, 1])
    with col1:
        new_item = st.text_input(
            "Nama item", key="new_item_input",
            label_visibility="collapsed", placeholder="Budi, Cuci baju, Ruangan 1, Film A..."
        )
    with col2:
        if st.button("Tambah", use_container_width=True) and new_item:
            add_item(new_item)
            st.rerun()

    st.subheader("Tambah sekaligus dari file")
    up_col1, up_col2 = st.columns(2)
    with up_col1:
        excel_file = st.file_uploader("Unggah Excel (.xlsx / .xls)", type=["xlsx", "xls"],
            help= "Setiap sel yang terisi akan dianggap sebagai satu item, termasuk dari semua sheet dan kolom."
        )
        handle_upload(excel_file, parse_excel, "excel")
    with up_col2:
        docx_file = st.file_uploader("Unggah Word (.docx)", type=["docx"],
            help= "Setiap paragraf atau baris yang dipisahkan dengan Enter akan dianggap sebagai satu item. Item yang sama akan otomatis dihapus."
        )
        handle_upload(docx_file, parse_docx, "docx")
    st.divider()
    st.subheader(f"Item saat ini ({len(st.session_state.item_list)})")
    if not st.session_state.item_list:
        st.info("Belum ada item. Tambahkan item baru di atas atau unggah file Excel/Word.")
    else:
        for item in list(st.session_state.item_list):
            c1, c2 = st.columns([5, 1])
            c1.write(item)
            if c2.button("Hapus", key=f"rm_{item}"):
                remove_item(item)
                st.rerun()

# --- Tab 3: Aturan ------------------------------------------------------
with tab_constraints:
    if not st.session_state.item_list:
        st.info("Tambah item terlebih dahulu di tab 'Item'.")
    else:
        st.subheader("Banyak item per hari")
        cap_mode = st.radio(
            "Cara menentukan jumlah item:", 
            [
                "Otomatis", 
                "Manual (Atur sendiri)"
            ], 
                horizontal=True,
            index=0 if st.session_state.capacity_mode == "Otomatis" else 1,
        )
        st.session_state.capacity_mode = cap_mode

        target_mode = st.radio(
            "Yang ingin dibuat sama rata:",
            [
                "Jumlah item per hari",
                "Kemunculan setiap item",
            ],
            horizontal=True,
            index=0 if st.session_state.capacity_target_mode == "Tepat" else 1,
            help=(
                "Jumlah item per hari: misalnya target 3 item per hari, maka setiap hari "
                "akan diisi tepat 3 item. Jumlah kemunculan tiap item bisa berbeda.\n\n"
                "Kemunculan setiap item: setiap item akan mendapat jumlah kemunculan "
                "yang sama. Karena pembagiannya harus merata, beberapa hari bisa berisi "
                "lebih banyak atau lebih sedikit item."
            ),
        )
        st.session_state.capacity_target_mode = "Tepat" if target_mode.startswith("Tepat") else "Maksimal"

        if cap_mode == "Manual (Atur sendiri)":
            st.session_state.capacity_manual = st.number_input(
                "Item per hari", min_value=1, max_value=100,
                value=st.session_state.capacity_manual,
            )
            preview_capacity = st.session_state.capacity_manual
        else:
            preview_capacity = compute_capacity(
                len(st.session_state.item_list), len(avail_day_keys), "Otomatis", 1
            )

        if len(avail_day_keys) < len(day_keys):
            st.caption(f"({len(day_keys) - len(avail_day_keys)} hari/tanggal dikecualikan sebagai hari libur.)")

        preview_occ = compute_occurrences(
            st.session_state.item_list, len(avail_day_keys), preview_capacity,
            st.session_state.capacity_target_mode,
        )
        judul_expander = (
            "Pratinjau: berapa kali setiap item akan muncul dalam seminggu?"
            if st.session_state.scope == "Seminggu"
            else "Pratinjau: berapa kali setiap item akan muncul dalam sebulan?"
        )

        with st.expander(judul_expander):
            st.table(pd.DataFrame(
                {"Item": list(preview_occ.keys()), "Kemunculan": list(preview_occ.values())}
            ))

        st.divider()
        if st.session_state.scope == "Seminggu":
            subheader2 = "Hari yang ingin dikosongkan"
            help2="Jadwal tidak akan dibuat pada hari yang dipilih."
        else:
            subheader2 = "Tanggal yang ingin dikosongkan"
            help2="Jadwal tidak akan dibuat pada tanggal yang dipilih."
        st.subheader(subheader2, help=help2)
        if st.session_state.scope == "Seminggu":
            selected_excl = st.multiselect(
                "Pilih hari libur", day_keys,
                default=[d for d in day_keys if d in st.session_state.excluded_days],
                key="excluded_days_select", placeholder="Pilih hari...",
            )
            st.session_state.excluded_days = set(selected_excl)
        else:
            month_start = dt.date(st.session_state.year, st.session_state.month, 1)
            month_end = dt.date(
                st.session_state.year, st.session_state.month,
                calendar.monthrange(st.session_state.year, st.session_state.month)[1],
            )
            picked_excl = st.date_input(
                "Pilih tanggal atau rentang tanggal libur:", value=(month_start, month_start),
                min_value=month_start, max_value=month_end, key="excl_day_pick", format="DD-MM-YYYY",
            )
            if st.button("Tambah Hari Libur", key="add_excl_day"):
                if isinstance(picked_excl, tuple) and len(picked_excl) == 2:
                    new_dates = dates_in_range(picked_excl[0], picked_excl[1])
                else:
                    single = picked_excl if isinstance(picked_excl, dt.date) else picked_excl[0]
                    new_dates = [single.isoformat()]
                st.session_state.excluded_days.update(new_dates)
                st.rerun()

            current_excl = sorted(st.session_state.excluded_days)
            if current_excl:
                for d in current_excl:
                    c1, c2 = st.columns([4, 1])
                    c1.write(label_of.get(d, d))
                    if c2.button("Hapus", key=f"rmexcl_{d}"):
                        st.session_state.excluded_days.discard(d)
                        st.rerun()
            else:
                st.caption("Belum ada hari atau tanggal libur.")

        st.divider()
        if st.session_state.scope == "Seminggu":
            subheader3 = "Hari yang dilarang untuk setiap item"
            help3 = "Pilih hari yang dilarang untuk item ini."
        else:
            subheader3 = "Tanggal yang dilarang untuk setiap item"
            help3 = "Tambahkan satu tanggal atau rentang tanggal yang dilarang untuk setiap item."
        st.subheader(subheader3, help=help3)
        if st.session_state.scope == "Seminggu":
            for item in st.session_state.item_list:
                current = st.session_state.banned_days.get(item, set())
                selected = st.multiselect(
                    item, day_keys, default=[d for d in day_keys if d in current], key=f"ban_{item}", placeholder="Pilih hari..."
                )
                st.session_state.banned_days[item] = set(selected)
        else:
            month_start = dt.date(st.session_state.year, st.session_state.month, 1)
            month_end = dt.date(
                st.session_state.year, st.session_state.month,
                calendar.monthrange(st.session_state.year, st.session_state.month)[1],
            )
            for item in st.session_state.item_list:
                with st.expander(f"{item} ({len(st.session_state.banned_days.get(item, set()))} hari dilarang)"):
                    picked = st.date_input(
                        "Pilih satu tanggal atau rentang tanggal.", value=(month_start, month_start),
                        min_value=month_start, max_value=month_end, key=f"banpick_{item}", format="DD-MM-YYYY"
                    )
                    if st.button("Tambah Aturan", key=f"addban_{item}"):
                        if isinstance(picked, tuple) and len(picked) == 2:
                            new_dates = dates_in_range(picked[0], picked[1])
                        else:
                            single = picked if isinstance(picked, dt.date) else picked[0]
                            new_dates = [single.isoformat()]
                        st.session_state.banned_days.setdefault(item, set()).update(new_dates)
                        st.rerun()

                    current = sorted(st.session_state.banned_days.get(item, set()))
                    hari = {
                        0: "Senin", 1: "Selasa", 2: "Rabu",
                        3: "Kamis", 4: "Jumat", 5: "Sabtu", 6: "Minggu"
                    }

                    if current:
                        for d in current:
                            c1, c2 = st.columns([4, 1])
                            c1.write(label_of.get(d, str(d)))
                            if c2.button("Hapus", key=f"rmban_{item}_{d}"):
                                st.session_state.banned_days[item].discard(d)
                                st.rerun()
                    else:
                        st.caption("Item ini bisa kapan saja.")

        st.divider()
        if st.session_state.scope == "Seminggu":
            subheader4 = "Item yang harus dijadwalkan di hari berbeda"
        else:
            subheader4 = "Item yang harus dijadwalkan di tanggal berbeda"
        st.subheader(subheader4)
        colA, colB, colC = st.columns([3, 3, 1])
        with colA:
            item_a = st.selectbox("Item A", st.session_state.item_list, key="excl_a")
        other_items = [i for i in st.session_state.item_list if i != item_a]
        with colB:
            item_b = st.selectbox("Item B", other_items, key="excl_b") if other_items else None
        with colC:
            st.write("")
            st.write("")
            if st.button("Tambah Aturan") and item_b:
                pair = tuple(sorted((item_a, item_b)))
                if pair not in st.session_state.exclusion_pairs:
                    st.session_state.exclusion_pairs.append(pair)
                st.rerun()

        if st.session_state.exclusion_pairs:
            for pair in list(st.session_state.exclusion_pairs):
                c1, c2 = st.columns([5, 1])
                c1.write(f"**{pair[0]}**  dan  **{pair[1]}**   harus dijadwalkan di hari atau tanggal berbeda.")
                if c2.button("Hapus", key=f"rmpair_{pair[0]}_{pair[1]}"):
                    st.session_state.exclusion_pairs.remove(pair)
                    st.rerun()
        else:
            st.caption("Belum ada item yang perlu dijadwalkan di hari atau tanggal berbeda.")

# --- Tab 4: Jadwalkan ----------------------------------------------------
with tab_run:
    items = st.session_state.item_list
    if not items:
        st.info("Tambah item terlebih dahulu di tab 'Item'.")
    elif not avail_day_keys:
        st.warning("Semua hari atau tanggal saat ini menjadi hari libur. Tidak ada yang bisa dijadwalkan.")
    else:
        capacity = compute_capacity(len(items), len(avail_day_keys), st.session_state.capacity_mode, st.session_state.capacity_manual)
        target_mode = st.session_state.capacity_target_mode
        occurrences = compute_occurrences(items, len(avail_day_keys), capacity, target_mode)
        tokens = build_tokens(items, occurrences)

        st.subheader("Pengaturan Jadwal",
                     help="Biasanya bisa langsung klik 'Buat Jadwal'. Kalau hasilnya belum sesuai, atur cara sistem mencari jadwal di pengaturan lanjutan."
                     )
        with st.expander("Pengaturan lanjutan"):
            st.caption("Pengaturan ini menentukan seberapa banyak dan seberapa bervariasi susunan jadwal yang akan dicoba.")
            c1, c2, c3 = st.columns(3)
            pop_size = c1.number_input("Banyak kemungkinan", 20, 1000, 150, 10, help="Makin besar, makin banyak jadwal dicoba. Waktu proses bertambah.")
            generations = c2.number_input("Jumlah percobaan", 50, 3000, 300, 50, help="Makin besar, makin lama mencari susunan yang sesuai.")
            mutation_rate = c3.slider("Variasi pencarian", 0.0, 1.0, 0.15, help="Seberapa sering sistem mengubah susunan jadwal untuk mencari alternatif.")
            st.caption(f"{len(items)} item × target {capacity} item/hari × {len(avail_day_keys)} hari tersedia → {len(tokens)} jadwal yang harus diisi.")

        c1, c2, c3 = st.columns([1, 1, 1])
        with c2:
            if st.button("Buat Jadwal", type="primary", use_container_width=True):
                with st.spinner("Sedang menyusun jadwal..."):
                    chromo, score, history = run_ga(
                        tokens, avail_day_keys, st.session_state.banned_days, st.session_state.exclusion_pairs,
                        capacity, int(pop_size), int(generations), float(mutation_rate), target_mode
                    )

                st.session_state.result = {
                    "chromo": chromo, "score": score, "history": history, "tokens": tokens,
                    "capacity": capacity, "target_mode": target_mode,
                    "avail_day_keys": avail_day_keys, "avail_day_labels": avail_day_labels,
                    "full_day_keys": day_keys, "full_day_labels": day_labels,
                    "excluded_days": set(st.session_state.excluded_days),
                }

        result = st.session_state.get("result")
        stale = (
            not result
            or result["full_day_keys"] != day_keys
            or result["avail_day_keys"] != avail_day_keys
            or result["excluded_days"] != st.session_state.excluded_days
        )

        if result and not stale:
            st.success("Jadwal berhasil dibuat!")

            avail_items = {k: [] for k in result["avail_day_keys"]}
            for idx, d in enumerate(result["chromo"]):
                avail_items[result["avail_day_keys"][d]].append(result["tokens"][idx])

            # Susun kolom per hari (termasuk hari libur, ditandai khusus) — dipakai
            # untuk ekspor Excel dengan header = nama hari/tanggal.
            schedule_cols = {}
            for k in result["full_day_keys"]:
                label = label_of.get(k, k)
                if k in result["excluded_days"]:
                    schedule_cols[label] = ["(Libur)"]
                else:
                    its = avail_items.get(k, [])
                    schedule_cols[label] = its if its else ["-"]

            max_len = max(len(v) for v in schedule_cols.values())
            for label in schedule_cols:
                schedule_cols[label] += [""] * (max_len - len(schedule_cols[label]))

            schedule_df = pd.DataFrame(schedule_cols)  # dipakai HANYA untuk ekspor Excel

            weeks = get_calendar_weeks(
                st.session_state.scope, st.session_state.year, st.session_state.month,
                avail_items, result["excluded_days"],
            )

            st.subheader("Jadwal")
            st.markdown(build_calendar_html(weeks), unsafe_allow_html=True)

            st.subheader("Pemeriksaan Aturan")
            violations = diagnose(
                result["tokens"], result["chromo"], result["avail_day_keys"], result["avail_day_labels"],
                st.session_state.banned_days, st.session_state.exclusion_pairs, result["capacity"],
                result.get("target_mode", "Tepat"),
            )
            for v in violations:
                st.markdown(v)

            with st.expander("Info lanjutan"):
                st.caption(f"Skor fitness akhir: **{result['score']}**")
                st.line_chart(pd.DataFrame(result["history"]).set_index("gen"))

            excel_buffer = BytesIO()
            schedule_df.to_excel(excel_buffer, index=False)
            pdf_bytes = build_calendar_pdf(
                weeks, st.session_state.scope, result["capacity"], result.get("target_mode", "Tepat"),
            )

            dl1, dl2 = st.columns(2)
            with dl1:
                st.download_button(
                    "DOWNLOAD JADWAL (Excel)", excel_buffer.getvalue(), "jadwal.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with dl2:
                st.download_button(
                    "DOWNLOAD JADWAL (PDF)", pdf_bytes, "jadwal.pdf", "application/pdf",
                    use_container_width=True,
                )

        elif result:
            st.info("Pengaturan periode, tanggal, atau hari libur berubah sejak jadwal terakhir dibuat. Klik **Buat Jadwal** lagi.")
