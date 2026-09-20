"""
All-Purpose Scheduler (Genetic Algorithm) — Weekly or Monthly
--------------------------------------------------------------
Schedule any set of items across either the 7 days of a week or the
calendar days of a specific month. Items can be banned from specific
days/dates (or date ranges, in monthly mode), and pairs of items can be
marked as "never on the same day." A per-day capacity (manual or
automatically computed) decides how many item-slots each day needs,
and items repeat across days as many times as necessary to fill that
capacity evenly.

Run with:  streamlit run scheduler_app.py
"""

import calendar
import datetime as dt
import math
import random
from io import BytesIO

import pandas as pd
import streamlit as st

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
# Day-label helpers
# ----------------------------------------------------------------------
def get_day_keys(scope: str, month: int = None, year: int = None):
    """Returns (keys, labels) — keys are used internally, labels are
    what's shown in the UI. Same order, same length."""
    if scope == "Weekly":
        return list(WEEKDAYS), list(WEEKDAYS)
    n_days = calendar.monthrange(year, month)[1]
    keys, labels = [], []
    for day in range(1, n_days + 1):
        date = dt.date(year, month, day)
        keys.append(date.isoformat())
        labels.append(date.strftime("%d %b (%a)"))
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
        "banned_days": {},        # item -> set(day keys)
        "exclusion_pairs": [],    # list of (item_a, item_b)
        "scope": "Weekly",
        "month": today.month,
        "year": today.year,
        "capacity_mode": "Automatic",
        "capacity_manual": 1,
        "result": None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


init_state()

st.set_page_config(page_title="GA Scheduler", layout="wide")
st.title("🧬 All-Purpose Scheduler")
st.caption(
    "Schedule any set of items across a week or a specific month using a "
    "genetic algorithm — respecting banned days/dates, 'never on the same "
    "day' rules, and a target number of items per day."
)


# ----------------------------------------------------------------------
# Item helpers
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
        st.error("python-docx isn't installed — add 'python-docx' to requirements.txt.")
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
    st.success(f"Found {len(names)} unique item(s) in the file, added {added} new one(s).")


# ----------------------------------------------------------------------
# Capacity & occurrences
# ----------------------------------------------------------------------
def compute_capacity(n_items, n_days, mode, manual_value):
    if n_items == 0 or n_days == 0:
        return 0
    if mode == "Manual":
        return max(1, int(manual_value))
    return max(1, math.ceil(n_items / n_days))


def compute_occurrences(items, n_days, capacity):
    n = len(items)
    if n == 0:
        return {}
    total_slots = capacity * n_days
    base, rem = divmod(total_slots, n)
    return {item: base + (1 if i < rem else 0) for i, item in enumerate(items)}


def build_tokens(items, occurrences):
    tokens = []
    for item in items:
        tokens.extend([item] * occurrences.get(item, 0))
    return tokens


# ----------------------------------------------------------------------
# Genetic algorithm
# ----------------------------------------------------------------------
def run_ga(tokens, day_keys, banned_days, exclusion_pairs, capacity,
           population_size=150, generations=300, mutation_rate=0.15):
    n_tokens = len(tokens)
    n_days = len(day_keys)
    day_pos = {k: i for i, k in enumerate(day_keys)}
    banned_idx = {
        item: {day_pos[d] for d in banned_days.get(item, set()) if d in day_pos}
        for item in set(tokens)
    }
    pair_list = list(exclusion_pairs)

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
            score -= abs(len(items_on_day) - capacity) * 300

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


def diagnose(tokens, chromo, day_keys, day_labels, banned_days, exclusion_pairs, capacity):
    violations = []
    day_items = [[] for _ in day_keys]
    for idx, d in enumerate(chromo):
        day_items[d].append(tokens[idx])

    for idx, d in enumerate(chromo):
        item = tokens[idx]
        if day_keys[d] in banned_days.get(item, set()):
            violations.append(f"🚨 **{item}** is scheduled on **{day_labels[d]}**, a banned day for it.")

    for i, items_on_day in enumerate(day_items):
        counts = {}
        for it in items_on_day:
            counts[it] = counts.get(it, 0) + 1
        for it, c in counts.items():
            if c > 1:
                violations.append(f"🚨 **{it}** appears **{c} times** on **{day_labels[i]}** (should be at most once).")

    day_sets = [set(x) for x in day_items]
    for a, b in exclusion_pairs:
        for i, s in enumerate(day_sets):
            if a in s and b in s:
                violations.append(f"⚠️ **{a}** and **{b}** are both on **{day_labels[i]}** (must be on different days).")

    for i, items_on_day in enumerate(day_items):
        if len(items_on_day) != capacity:
            violations.append(
                f"📉 **{day_labels[i]}** has **{len(items_on_day)}** item(s) scheduled, "
                f"target is **{capacity}**."
            )

    if not violations:
        violations.append("✅ No constraint violations — every rule is satisfied.")
    return violations


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Reset")
    if st.button("Clear everything", use_container_width=True):
        for key in ("item_list", "exclusion_pairs"):
            st.session_state[key] = []
        st.session_state["banned_days"] = {}
        st.session_state["result"] = None
        st.rerun()

tab_scope, tab_items, tab_constraints, tab_run = st.tabs(
    ["0️⃣ Scope", "1️⃣ Items", "2️⃣ Constraints", "3️⃣ Run & Results"]
)

# --- Tab 0: Scope ---------------------------------------------------------
with tab_scope:
    st.subheader("What are you scheduling for?")
    scope = st.radio("Scope", ["Weekly", "Monthly"], horizontal=True,
                      index=0 if st.session_state.scope == "Weekly" else 1)
    if scope != st.session_state.scope:
        st.session_state.scope = scope
        st.session_state.result = None  # day-count changed, old result is stale

    if scope == "Monthly":
        c1, c2 = st.columns(2)
        month = c1.selectbox("Month", list(range(1, 13)),
                              format_func=lambda m: MONTH_NAMES[m - 1],
                              index=st.session_state.month - 1)
        year = c2.number_input("Year", min_value=2000, max_value=2100, value=st.session_state.year)
        if month != st.session_state.month or year != st.session_state.year:
            st.session_state.month = month
            st.session_state.year = year
            st.session_state.result = None
        n_days = calendar.monthrange(int(year), int(month))[1]
        st.caption(f"{MONTH_NAMES[month - 1]} {year} has {n_days} days.")
    else:
        st.caption("Using the 7 fixed days of the week.")

day_keys, day_labels = get_day_keys(
    st.session_state.scope, st.session_state.month, st.session_state.year
)
label_of = dict(zip(day_keys, day_labels))

# --- Tab 1: Items ----------------------------------------------------------
with tab_items:
    st.subheader("Add an item")
    col1, col2 = st.columns([4, 1])
    with col1:
        new_item = st.text_input(
            "Item name", key="new_item_input",
            label_visibility="collapsed", placeholder="e.g. Alice, Task A, Room 1..."
        )
    with col2:
        if st.button("Add", use_container_width=True) and new_item:
            add_item(new_item)
            st.rerun()

    st.subheader("Bulk import")
    up_col1, up_col2 = st.columns(2)
    with up_col1:
        excel_file = st.file_uploader("Upload Excel (.xlsx / .xls)", type=["xlsx", "xls"])
        handle_upload(excel_file, parse_excel, "excel")
    with up_col2:
        docx_file = st.file_uploader("Upload Word (.docx)", type=["docx"])
        handle_upload(docx_file, parse_docx, "docx")
    st.caption(
        "Excel: every non-empty cell across all sheets/columns is treated as one item. "
        "Word: every non-empty paragraph (each Enter-separated line) is one item. "
        "Duplicates are dropped automatically either way."
    )

    st.divider()
    st.subheader(f"Current items ({len(st.session_state.item_list)})")
    if not st.session_state.item_list:
        st.info("No items yet — add some above.")
    else:
        for item in list(st.session_state.item_list):
            c1, c2 = st.columns([5, 1])
            c1.write(item)
            if c2.button("Remove", key=f"rm_{item}"):
                remove_item(item)
                st.rerun()

# --- Tab 2: Constraints ------------------------------------------------------
with tab_constraints:
    if not st.session_state.item_list:
        st.info("Add items first in the 'Items' tab.")
    else:
        st.subheader("Capacity — how many items per day")
        cap_mode = st.radio(
            "Mode", ["Automatic", "Manual"], horizontal=True,
            index=0 if st.session_state.capacity_mode == "Automatic" else 1,
        )
        st.session_state.capacity_mode = cap_mode
        if cap_mode == "Manual":
            st.session_state.capacity_manual = st.number_input(
                "Items per day", min_value=1, max_value=max(1, len(st.session_state.item_list)),
                value=st.session_state.capacity_manual,
            )
            preview_capacity = st.session_state.capacity_manual
        else:
            preview_capacity = compute_capacity(
                len(st.session_state.item_list), len(day_keys), "Automatic", 1
            )
            st.caption(f"Automatic capacity for {len(st.session_state.item_list)} item(s) across "
                       f"{len(day_keys)} day(s): **{preview_capacity} item(s)/day**.")

        preview_occ = compute_occurrences(st.session_state.item_list, len(day_keys), preview_capacity)
        with st.expander("Preview: how many times will each item appear?"):
            st.table(pd.DataFrame(
                {"Item": list(preview_occ.keys()), "Occurrences": list(preview_occ.values())}
            ))

        st.divider()
        st.subheader("Banned days / dates per item")
        if st.session_state.scope == "Weekly":
            st.caption("Pick the weekdays an item is NOT allowed to be scheduled on.")
            for item in st.session_state.item_list:
                current = st.session_state.banned_days.get(item, set())
                selected = st.multiselect(
                    item, day_keys, default=[d for d in day_keys if d in current], key=f"ban_{item}"
                )
                st.session_state.banned_days[item] = set(selected)
        else:
            st.caption("Add a single banned date, or a whole range, per item.")
            month_start = dt.date(st.session_state.year, st.session_state.month, 1)
            month_end = dt.date(
                st.session_state.year, st.session_state.month,
                calendar.monthrange(st.session_state.year, st.session_state.month)[1],
            )
            for item in st.session_state.item_list:
                with st.expander(f"{item} — {len(st.session_state.banned_days.get(item, set()))} banned date(s)"):
                    picked = st.date_input(
                        "Pick a date, or a start+end range", value=(month_start, month_start),
                        min_value=month_start, max_value=month_end, key=f"banpick_{item}",
                    )
                    if st.button("Add ban", key=f"addban_{item}"):
                        if isinstance(picked, tuple) and len(picked) == 2:
                            new_dates = dates_in_range(picked[0], picked[1])
                        else:
                            single = picked if isinstance(picked, dt.date) else picked[0]
                            new_dates = [single.isoformat()]
                        st.session_state.banned_days.setdefault(item, set()).update(new_dates)
                        st.rerun()

                    current = sorted(st.session_state.banned_days.get(item, set()))
                    if current:
                        for d in current:
                            c1, c2 = st.columns([4, 1])
                            c1.write(label_of.get(d, d))
                            if c2.button("Remove", key=f"rmban_{item}_{d}"):
                                st.session_state.banned_days[item].discard(d)
                                st.rerun()
                    else:
                        st.caption("No banned dates for this item.")

        st.divider()
        st.subheader("Items that must NEVER share a day")
        colA, colB, colC = st.columns([3, 3, 1])
        with colA:
            item_a = st.selectbox("Item A", st.session_state.item_list, key="excl_a")
        other_items = [i for i in st.session_state.item_list if i != item_a]
        with colB:
            item_b = st.selectbox("Item B", other_items, key="excl_b") if other_items else None
        with colC:
            st.write("")
            st.write("")
            if st.button("Add rule") and item_b:
                pair = tuple(sorted((item_a, item_b)))
                if pair not in st.session_state.exclusion_pairs:
                    st.session_state.exclusion_pairs.append(pair)
                st.rerun()

        if st.session_state.exclusion_pairs:
            for pair in list(st.session_state.exclusion_pairs):
                c1, c2 = st.columns([5, 1])
                c1.write(f"{pair[0]}  ⟷  {pair[1]}   (never same day)")
                if c2.button("Remove", key=f"rmpair_{pair[0]}_{pair[1]}"):
                    st.session_state.exclusion_pairs.remove(pair)
                    st.rerun()
        else:
            st.caption("No exclusion rules yet.")

# --- Tab 3: Run & Results ----------------------------------------------------
with tab_run:
    if not st.session_state.item_list:
        st.info("Add items first in the 'Items' tab.")
    else:
        capacity = compute_capacity(
            len(st.session_state.item_list), len(day_keys),
            st.session_state.capacity_mode, st.session_state.capacity_manual,
        )
        occurrences = compute_occurrences(st.session_state.item_list, len(day_keys), capacity)
        tokens = build_tokens(st.session_state.item_list, occurrences)

        st.subheader("GA parameters")
        c1, c2, c3 = st.columns(3)
        pop_size = c1.number_input("Population size", 20, 1000, 150, step=10)
        generations = c2.number_input("Generations", 50, 3000, 300, step=50)
        mutation_rate = c3.slider("Mutation rate", 0.0, 1.0, 0.15)
        st.caption(f"{len(st.session_state.item_list)} item(s) × target {capacity}/day × "
                   f"{len(day_keys)} day(s) → {len(tokens)} total slots to fill.")

        if st.button("🚀 Generate Schedule", type="primary"):
            with st.spinner("Running genetic algorithm..."):
                chromo, score, history = run_ga(
                    tokens, day_keys, st.session_state.banned_days,
                    st.session_state.exclusion_pairs, capacity,
                    population_size=int(pop_size), generations=int(generations),
                    mutation_rate=float(mutation_rate),
                )
            st.session_state.result = {
                "chromo": chromo, "score": score, "history": history,
                "tokens": tokens, "capacity": capacity,
                "day_keys": day_keys, "day_labels": day_labels,
            }

        result = st.session_state.result
        if result and result["day_keys"] == day_keys:
            st.metric("Final fitness score", result["score"])

            schedule = {k: [] for k in result["day_keys"]}
            for idx, d in enumerate(result["chromo"]):
                schedule[result["day_keys"][d]].append(result["tokens"][idx])

            st.subheader("Schedule")
            n_cols = min(7, len(result["day_keys"]))
            for start in range(0, len(result["day_keys"]), n_cols):
                cols = st.columns(n_cols)
                chunk = result["day_keys"][start:start + n_cols]
                for c, k in zip(cols, chunk):
                    with c:
                        st.markdown(f"**{label_of.get(k, k)}**")
                        if schedule[k]:
                            for it in schedule[k]:
                                st.write(f"- {it}")
                        else:
                            st.caption("—")

            st.subheader("Constraint report")
            for v in diagnose(
                result["tokens"], result["chromo"], result["day_keys"], result["day_labels"],
                st.session_state.banned_days, st.session_state.exclusion_pairs, result["capacity"],
            ):
                st.markdown(v)

            st.subheader("Fitness evolution")
            hist_df = pd.DataFrame(result["history"]).set_index("gen")
            st.line_chart(hist_df)

            export_df = pd.DataFrame(
                {"Day": label_of.get(result["day_keys"][d], result["day_keys"][d]), "Item": result["tokens"][idx]}
                for idx, d in enumerate(result["chromo"])
            ).sort_values("Day")
            buffer = BytesIO()
            export_df.to_excel(buffer, index=False)
            st.download_button(
                "⬇️ Download schedule (Excel)",
                data=buffer.getvalue(),
                file_name="schedule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        elif result:
            st.info("Scope/date settings changed since the last run — click 'Generate Schedule' again.")
