import streamlit as st
import pdfplumber
import pandas as pd
import re
import tempfile
import base64
import hashlib
import difflib

# =============================
# 🎯 CONFIG
# =============================
st.set_page_config(page_title="Smart PDF Comparator", layout="wide")

# =============================
# 🔐 LOGIN SYSTEM
# =============================
def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        st.markdown("""
        <div style="display:flex; justify-content:center; margin-top:80px;">
        <div style="background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
                    padding: 2.5rem; border-radius: 20px; width:380px; text-align:center;">
            <h1 style="color:#e94560; margin-bottom:0.2rem;">🏭 Smart PDF</h1>
            <h2 style="color:#e94560; margin-top:0;">Comparator</h2>
            <p style="color:#a8b2d8; margin-bottom:2rem;">Please login to continue</p>
        </div>
        </div>
        """, unsafe_allow_html=True)

        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            username = st.text_input("👤 Username")
            password = st.text_input("🔑 Password", type="password")

            if st.button("🔓 Login", use_container_width=True):
                if username == "dreamy" and password == "YDreamy":
                    st.session_state.logged_in = True
                    st.rerun()
                else:
                    st.error("❌ Incorrect username or password")
        st.stop()

check_login()

# =============================
# 🧠 DEBUG MODE
# =============================
st.sidebar.title("⚙️ Settings")
debug_mode = st.sidebar.checkbox("🐞 Debug mode", value=False)

if st.sidebar.button("🚪 Logout"):
    st.session_state.logged_in = False
    st.rerun()

# =============================
# 📄 VIEW PDF
# =============================
def show_pdf(file):
    file.seek(0)
    base64_pdf = base64.b64encode(file.read()).decode("utf-8")
    pdf_display = f"""
    <embed src="data:application/pdf;base64,{base64_pdf}" 
    width="100%" height="600px" type="application/pdf">
    """
    st.markdown(pdf_display, unsafe_allow_html=True)

# =============================
# 📑 INDEX
# =============================
def extract_index_and_info(pdf_path):
    index_map = {}
    project_name = "Projet inconnu"

    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 2:
            return {}, "", project_name

        text = pdf.pages[1].extract_text() or ""

        for line in text.split("\n"):
            if "project name" in line.lower():
                project_name = line.replace("Project Name", "").strip()

            if re.match(r"^\s*\d+\s+.+", line):
                match = re.match(r"^(\d+)\s+(.+)", line)
                if match:
                    index_map[match.group(2)] = int(match.group(1))

    return index_map, text, project_name

# =============================
# 🔥 FILTER INDEX
# =============================
def filter_sections(index_map):
    return {
        k: v for k, v in index_map.items()
        if "cover" not in k.lower() and "index" not in k.lower()
    }

# =============================
# 🔥 PAGE RANGE
# =============================
def get_section_ranges(index_map, total_pages):
    sorted_items = sorted(index_map.items(), key=lambda x: x[1])
    ranges = {}

    for i, (section, start) in enumerate(sorted_items):
        if i < len(sorted_items) - 1:
            end = sorted_items[i + 1][1] - 1
        else:
            end = total_pages

        ranges[section] = (start, end)

    return ranges

# =============================
# ⚡ PAGE DATA
# =============================
def get_pages_data(pdf_path):
    pages = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            clean = re.sub(r"\s+", " ", text).strip().lower()
            h = hashlib.md5(clean.encode()).hexdigest()
            pages.append((clean, h))

    return pages

# =============================
# ⚡ QUICK COMPARE
# =============================
def quick_compare(index_new, index_old, path_new, path_old):

    new_filtered = filter_sections(index_new)
    old_filtered = filter_sections(index_old)

    if set(new_filtered.keys()) != set(old_filtered.keys()):
        return False, {"reason": "Different structure (sections)"}

    for k in new_filtered:
        if new_filtered[k] != old_filtered[k]:
            return False, {"reason": f"Different page mapping in section: {k}"}

    new_pages = get_pages_data(path_new)
    old_pages = get_pages_data(path_old)

    min_len = min(len(new_pages), len(old_pages))
    diff_pages = []
    scores = []

    for i in range(min_len):
        txt1 = new_pages[i][0]
        txt2 = old_pages[i][0]
        ratio = difflib.SequenceMatcher(None, txt1, txt2).ratio()
        scores.append(ratio)
        if ratio < 0.99:
            diff_pages.append(i + 1)

    extra_pages = abs(len(new_pages) - len(old_pages))
    if extra_pages > 0:
        diff_pages.extend(list(range(min_len + 1, min_len + extra_pages + 1)))

    avg_similarity = sum(scores) / len(scores) if scores else 0
    similarity = round(avg_similarity * 100, 2)
    difference = round(100 - similarity, 2)

    if len(diff_pages) == 0:
        return True, {"reason": "Fully identical", "similarity": 100, "difference": 0}

    return False, {
        "reason": "Same structure but DIFFERENT CONTENT",
        "similarity": similarity,
        "difference": difference,
        "different_pages": diff_pages[:10]
    }

# =============================
# 🧠 CLEAN META ROWS
# =============================
def is_meta_row(row):
    text = " ".join([str(x) for x in row if pd.notna(x)]).lower()
    keywords = ["project", "client", "rev", "date", "page", "document", "sheet", "title"]
    return any(k in text for k in keywords)

# =============================
# 🧠 CLEAN COLUMNS
# =============================
def clean_columns(cols):
    seen = {}
    clean = []

    for c in cols:
        c = str(c).strip()
        if c == "" or c.lower() == "nan":
            c = "col"
        if c in seen:
            seen[c] += 1
            clean.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            clean.append(c)

    return clean

# =============================
# 🔑 HEADER DETECTION
# =============================
def detect_real_header(pdf_path, section_start_page, nb_cols):
    with pdfplumber.open(pdf_path) as pdf:
        page_idx = section_start_page - 1
        if page_idx >= len(pdf.pages):
            return None

        tables = pdf.pages[page_idx].extract_tables()
        if not tables:
            return None

        rows = [[str(c).strip() if c else "" for c in row] for row in tables[0]]

        if len(rows) < 3:
            return None

        data_start = None
        for i, r in enumerate(rows):
            if re.match(r"^\d+$", str(r[0]).strip()):
                data_start = i
                break

        if data_start is None or data_start < 2:
            return None

        h1 = rows[data_start - 2]
        h2 = rows[data_start - 1]

        max_len = max(len(h1), len(h2), nb_cols)
        h1 += [""] * (max_len - len(h1))
        h2 += [""] * (max_len - len(h2))

        merged = []
        for a, b in zip(h1[:nb_cols], h2[:nb_cols]):
            a, b = a.strip(), b.strip()
            merged.append(f"{a} {b}".strip() if a and b else a or b or "col")

        return merged

# =============================
# 📊 EXTRACT TABLES
# =============================
def extract_tables_range(pdf_path, start, end):
    tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for i in range(start - 1, end):
            for t in pdf.pages[i].extract_tables():
                df = pd.DataFrame(t)
                df.dropna(how="all", inplace=True)
                df = df.map(lambda x: str(x).strip() if pd.notna(x) else x)
                tables.append(df)

    return tables

# =============================
# 🧹 MERGE CLEAN
# =============================
def merge_tables_clean(tables, pdf_path, section_start_page):

    if not tables:
        return pd.DataFrame()

    nb_cols = next((len(df.columns) for df in tables if not df.empty), None)
    if nb_cols is None:
        return pd.DataFrame()

    header = detect_real_header(pdf_path, section_start_page, nb_cols)
    header = clean_columns(header) if header else [f"Col{i+1}" for i in range(nb_cols)]

    final = []

    for df in tables:
        df = df.dropna(how="all").reset_index(drop=True)
        if df.empty:
            continue

        df = df[~df.apply(is_meta_row, axis=1)]
        df = df[df.iloc[:, 0].astype(str).str.match(r"^\d+$", na=False)]

        if df.empty:
            continue

        if len(df.columns) < len(header):
            for i in range(len(header) - len(df.columns)):
                df[f"_extra_{i}"] = ""
        else:
            df = df.iloc[:, :len(header)]

        df.columns = header
        final.append(df)

    return pd.concat(final, ignore_index=True) if final else pd.DataFrame()

# =============================
# 🔬 DEEP COMPARISON ENGINE
# =============================
def deep_compare(df_new, df_old, key_col, selected_cols):

    df_new = df_new.fillna("")
    df_old = df_old.fillna("")

    df_new = df_new.set_index(key_col)
    df_old = df_old.set_index(key_col)

    new_keys = set(df_new.index)
    old_keys = set(df_old.index)

    added   = list(new_keys - old_keys)
    deleted = list(old_keys - new_keys)
    common  = list(new_keys & old_keys)

    modified = {}

    for k in common:
        changes = {}
        for col in selected_cols:
            if col in df_new.columns and col in df_old.columns:
                v1 = str(df_old.loc[k, col])
                v2 = str(df_new.loc[k, col])
                if v1 != v2:
                    changes[col] = {"old": v1, "new": v2}
        if changes:
            modified[k] = changes

    return {"added": added, "deleted": deleted, "modified": modified}

# =============================
# 🎨 DISPLAY COMPARISON RESULTS
# =============================
def display_comparison_results(result, section):

    added    = result["added"]
    deleted  = result["deleted"]
    modified = result["modified"]
    total    = len(added) + len(deleted) + len(modified)

    st.markdown("---")
    st.subheader("📊 KPI Summary")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("➕ Added",         len(added),
              delta=f"+{len(added)}" if added else None,
              delta_color="normal")
    c2.metric("➖ Deleted",       len(deleted),
              delta=f"-{len(deleted)}" if deleted else None,
              delta_color="inverse")
    c3.metric("✏️ Modified",      len(modified))
    c4.metric("🔢 Total changes", total)

    st.markdown("---")

    with st.expander(f"➕ Added equipment — {len(added)} item(s)",
                     expanded=len(added) > 0):
        if added:
            df_add = pd.DataFrame({"#": range(1, len(added)+1),
                                   "Equipment": sorted(added)})
            st.dataframe(df_add, use_container_width=True, hide_index=True)
        else:
            st.success("No equipment added.")

    with st.expander(f"➖ Deleted equipment — {len(deleted)} item(s)",
                     expanded=len(deleted) > 0):
        if deleted:
            df_del = pd.DataFrame({"#": range(1, len(deleted)+1),
                                   "Equipment": sorted(deleted)})
            st.dataframe(df_del, use_container_width=True, hide_index=True)
        else:
            st.success("No equipment deleted.")

    with st.expander(f"✏️ Modified parameters — {len(modified)} equipment",
                     expanded=True):
        if not modified:
            st.success("No parameter changes detected.")
            return

        rows = []
        for equip, changes in modified.items():
            for param, vals in changes.items():
                rows.append({
                    "Equipment": equip.replace("\n", " ").strip(),
                    "Parameter": param.replace("\n", " ").strip(),
                    "Old Value": vals["old"],
                    "New Value": vals["new"],
                })

        df_mod = (pd.DataFrame(rows)
                    .sort_values(["Equipment", "Parameter"])
                    .reset_index(drop=True))
        df_mod.index += 1

        st.session_state["df_mod"] = df_mod
        st.session_state["section"] = section

        def color_row(row):
            return [
                "",
                "",
                "background-color:#fff3cd; color:#7d5a00",
                "background-color:#d4edda; color:#155724",
            ]

        st.markdown("#### 🔍 Filter by parameter")
        all_params = sorted(df_mod["Parameter"].unique())
        chosen = st.multiselect(
            "Select parameter(s) to isolate",
            all_params,
            default=[],
            key="param_filter"
        )

        df_display = df_mod.copy()
        if chosen:
            df_display = df_mod[df_mod["Parameter"].isin(chosen)].reset_index(drop=True)
            df_display.index += 1

        st.dataframe(
            df_display.style.apply(color_row, axis=1),
            use_container_width=True,
            hide_index=False,
            height=min(600, 40 + len(df_display) * 36),
        )

        st.caption(f"Showing {len(df_display)} of {len(df_mod)} changes")

        csv = df_mod.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="⬇️ Download all changes as CSV",
            data=csv,
            file_name=f"changes_{section.replace(' ', '_')}.csv",
            mime="text/csv",
        )


# =============================
# 🧭 UI
# =============================
st.sidebar.title("📌 Menu")
menu = st.sidebar.radio("", [
    "🏠 Accueil",
    "⚡ Quick Compare",
    "📄 Viewer",
    "📑 Index",
    "📊 Tables",
    "🔬 Deep Comparison"
])

new_pdf = st.sidebar.file_uploader("🆕 New List", type=["pdf"])
old_pdf = st.sidebar.file_uploader("📁 Old List", type=["pdf"])

# =============================
# 🏠 HOME — DASHBOARD
# =============================
if menu == "🏠 Accueil":

    st.markdown("""
    <div style="background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
                padding: 2rem 2.5rem; border-radius: 16px; margin-bottom: 1.5rem;">
        <h1 style="color:#e94560; margin:0; font-size:2.2rem;">🏭 Smart PDF Comparator</h1>
        <p style="color:#a8b2d8; margin:0.4rem 0 0; font-size:1rem;">
            Industrial Electrical Load List — Comparison & Analysis Platform
        </p>
        <p style="color:#6c757d; margin:0.6rem 0 0; font-size:0.82rem;">
            Prepared by <strong style="color:#e94560;">Khadija Hodar</strong>
        </p>
    </div>
    """, unsafe_allow_html=True)

    deep_result  = st.session_state.get("deep_result", None)
    deep_section = st.session_state.get("deep_section", "—")

    has_result = deep_result is not None
    n_added    = len(deep_result["added"])    if has_result else None
    n_deleted  = len(deep_result["deleted"])  if has_result else None
    n_modified = len(deep_result["modified"]) if has_result else None
    n_total    = (n_added + n_deleted + n_modified) if has_result else None

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        if new_pdf:
            st.success(f"🆕 NEW PDF loaded — `{new_pdf.name}`")
        else:
            st.warning("🆕 NEW PDF not uploaded yet")
    with col_s2:
        if old_pdf:
            st.success(f"📁 OLD PDF loaded — `{old_pdf.name}`")
        else:
            st.warning("📁 OLD PDF not uploaded yet")

    st.markdown("---")

    st.subheader("📊 Last Deep Comparison Results")

    if has_result:
        st.caption(f"Section analysed : **{deep_section}**")
        k1, k2, k3, k4 = st.columns(4)

        def kpi_card(col, emoji, label, value, bg, fg):
            col.markdown(f"""
            <div style="background:{bg}; border-radius:12px; padding:1.2rem 1rem;
                        text-align:center; border:1px solid {fg}33;">
                <div style="font-size:2rem;">{emoji}</div>
                <div style="font-size:2rem; font-weight:700; color:{fg};">{value}</div>
                <div style="font-size:0.85rem; color:#888; margin-top:4px;">{label}</div>
            </div>
            """, unsafe_allow_html=True)

        kpi_card(k1, "➕", "Added",         n_added,    "#eafaf1", "#1e8449")
        kpi_card(k2, "➖", "Deleted",       n_deleted,  "#fdf2f2", "#c0392b")
        kpi_card(k3, "✏️", "Modified",      n_modified, "#fef9e7", "#d68910")
        kpi_card(k4, "🔢", "Total Changes", n_total,    "#eaf2fb", "#1a5276")

        st.markdown("")

        import json

        modified = deep_result["modified"]
        rows_mod = []
        for equip, changes in modified.items():
            for param, vals in changes.items():
                rows_mod.append({
                    "Equipment": equip.replace("\n", " ").strip(),
                    "Parameter": param.replace("\n", " ").strip(),
                    "Old": vals["old"],
                    "New": vals["new"],
                })
        df_dash = pd.DataFrame(rows_mod) if rows_mod else pd.DataFrame()

        if not df_dash.empty:
            chart_col1, chart_col2 = st.columns(2)

            with chart_col1:
                st.markdown("#### 🍩 Change breakdown")
                donut_labels = ["Added", "Deleted", "Modified"]
                donut_values = [n_added, n_deleted, n_modified]
                donut_colors = ["#1e8449", "#c0392b", "#d68910"]

                donut_html = f"""
                <canvas id="donut" width="280" height="280"></canvas>
                <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
                <script>
                new Chart(document.getElementById('donut'), {{
                    type: 'doughnut',
                    data: {{
                        labels: {json.dumps(donut_labels)},
                        datasets: [{{
                            data: {json.dumps(donut_values)},
                            backgroundColor: {json.dumps(donut_colors)},
                            borderWidth: 2,
                            borderColor: '#fff'
                        }}]
                    }},
                    options: {{
                        cutout: '60%',
                        plugins: {{
                            legend: {{ position: 'bottom' }},
                            tooltip: {{ callbacks: {{
                                label: ctx => ` ${{ctx.label}}: ${{ctx.parsed}}`
                            }}}}
                        }}
                    }}
                }});
                </script>
                """
                st.components.v1.html(donut_html, height=300)

            with chart_col2:
                st.markdown("#### 📊 Top modified parameters")
                param_counts = df_dash["Parameter"].value_counts().head(10)
                bar_labels   = param_counts.index.tolist()
                bar_values   = param_counts.values.tolist()

                bar_html = f"""
                <canvas id="bar" width="380" height="280"></canvas>
                <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
                <script>
                new Chart(document.getElementById('bar'), {{
                    type: 'bar',
                    data: {{
                        labels: {json.dumps(bar_labels)},
                        datasets: [{{
                            label: 'Changes',
                            data: {json.dumps(bar_values)},
                            backgroundColor: '#0f3460cc',
                            borderColor: '#e94560',
                            borderWidth: 1,
                            borderRadius: 6
                        }}]
                    }},
                    options: {{
                        indexAxis: 'y',
                        plugins: {{ legend: {{ display: false }} }},
                        scales: {{
                            x: {{ beginAtZero: true, ticks: {{ stepSize: 1 }} }},
                            y: {{ ticks: {{ font: {{ size: 11 }} }} }}
                        }}
                    }}
                }});
                </script>
                """
                st.components.v1.html(bar_html, height=300)

            st.markdown("#### 🏆 Top 5 most impacted equipment")
            top5 = (df_dash.groupby("Equipment")
                           .size()
                           .sort_values(ascending=False)
                           .head(5)
                           .reset_index())
            top5.columns = ["Equipment", "# Parameters changed"]
            top5.index += 1

            def highlight_top(row):
                colors = ["#fff9c4","#fff3cd","#fdebd0","#fdf2f8","#f0f4ff"]
                idx = row.name - 1
                c = colors[idx] if idx < len(colors) else ""
                return [f"background-color:{c}"]*len(row)

            st.dataframe(
                top5.style.apply(highlight_top, axis=1),
                use_container_width=True,
                hide_index=False,
            )

    else:
        st.info("💡 No comparison run yet. Go to **🔬 Deep Comparison**, run an analysis, then come back here to see the dashboard.")

        st.markdown("")
        k1, k2, k3, k4 = st.columns(4)
        for col, emoji, label in [
            (k1, "➕", "Added"),
            (k2, "➖", "Deleted"),
            (k3, "✏️", "Modified"),
            (k4, "🔢", "Total"),
        ]:
            col.markdown(f"""
            <div style="background:#f8f9fa; border-radius:12px; padding:1.2rem 1rem;
                        text-align:center; border:1px solid #dee2e6;">
                <div style="font-size:2rem;">{emoji}</div>
                <div style="font-size:2rem; font-weight:700; color:#ced4da;">—</div>
                <div style="font-size:0.85rem; color:#adb5bd; margin-top:4px;">{label}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div style="text-align:center; color:#6c757d; font-size:0.8rem; padding:0.5rem 0;">
        🏭 Smart PDF Comparator &nbsp;|&nbsp;
        Prepared by <strong>Khadija Hodar</strong> &nbsp;|&nbsp;
        Powered by pdfplumber · pandas · Streamlit
    </div>
    """, unsafe_allow_html=True)

# =============================
# ⚡ QUICK COMPARE
# =============================
elif menu == "⚡ Quick Compare":
    st.title("⚡ Quick Comparison")

    if new_pdf and old_pdf:

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as t1:
            t1.write(new_pdf.read())
            p1 = t1.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as t2:
            t2.write(old_pdf.read())
            p2 = t2.name

        idx_new, _, _ = extract_index_and_info(p1)
        idx_old, _, _ = extract_index_and_info(p2)

        same, info = quick_compare(idx_new, idx_old, p1, p2)

        if same:
            st.success("✅ PDFs are IDENTICAL")
            st.metric("Similarity (%)", info["similarity"])
            st.metric("Difference (%)", info["difference"])
        else:
            st.error("❌ PDFs are DIFFERENT")
            st.write("📌 Reason:", info["reason"])
            if "different_pages" in info:
                st.write("📄 Pages:", info["different_pages"])

# =============================
# 📄 VIEWER
# =============================
elif menu == "📄 Viewer":
    col1, col2 = st.columns(2)

    if new_pdf:
        with col1:
            st.subheader("🆕 NEW")
            show_pdf(new_pdf)

    if old_pdf:
        with col2:
            st.subheader("📁 OLD")
            show_pdf(old_pdf)

# =============================
# 📑 INDEX
# =============================
elif menu == "📑 Index":

    if new_pdf:
        with tempfile.NamedTemporaryFile(delete=False) as t:
            t.write(new_pdf.read())
            p = t.name
        idx, _, name = extract_index_and_info(p)
        st.title(name)
        st.write(filter_sections(idx))

    if old_pdf:
        with tempfile.NamedTemporaryFile(delete=False) as t:
            t.write(old_pdf.read())
            p = t.name
        idx, _, _ = extract_index_and_info(p)
        st.write(filter_sections(idx))

# =============================
# 📊 TABLES
# =============================
elif menu == "📊 Tables":

    st.title("📊 Deep Comparison (Tables)")

    if new_pdf and old_pdf:

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as t1:
            t1.write(new_pdf.read())
            p1 = t1.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as t2:
            t2.write(old_pdf.read())
            p2 = t2.name

        idx_new, _, _ = extract_index_and_info(p1)
        idx_old, _, _ = extract_index_and_info(p2)

        common = list(set(filter_sections(idx_new)) & set(filter_sections(idx_old)))

        if common:
            section = st.selectbox("🎯 Section", common)

            with pdfplumber.open(p1) as pdf:
                n1 = len(pdf.pages)
            with pdfplumber.open(p2) as pdf:
                n2 = len(pdf.pages)

            r1 = get_section_ranges(filter_sections(idx_new), n1)
            r2 = get_section_ranges(filter_sections(idx_old), n2)

            s1, e1 = r1[section]
            s2, e2 = r2[section]

            tnew = extract_tables_range(p1, s1, e1)
            told = extract_tables_range(p2, s2, e2)

            fnew = merge_tables_clean(tnew, p1, s1)
            fold = merge_tables_clean(told, p2, s2)

            col1, col2 = st.columns(2)

            with col1:
                st.subheader("🆕 NEW")
                st.dataframe(fnew)

            with col2:
                st.subheader("📁 OLD")
                st.dataframe(fold)

# =============================
# 🔬 DEEP COMPARISON UI
# =============================
elif menu == "🔬 Deep Comparison":

    st.title("🔬 Deep Column Comparison")

    if "deep_result" not in st.session_state:
        st.session_state["deep_result"] = None
    if "deep_section" not in st.session_state:
        st.session_state["deep_section"] = None

    if new_pdf and old_pdf:

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as t1:
            t1.write(new_pdf.read())
            p1 = t1.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as t2:
            t2.write(old_pdf.read())
            p2 = t2.name

        idx_new, _, _ = extract_index_and_info(p1)
        idx_old, _, _ = extract_index_and_info(p2)

        common = list(set(filter_sections(idx_new)) & set(filter_sections(idx_old)))

        if common:

            section = st.selectbox("🎯 Section", common)

            with pdfplumber.open(p1) as pdf:
                n1 = len(pdf.pages)
            with pdfplumber.open(p2) as pdf:
                n2 = len(pdf.pages)

            r1 = get_section_ranges(filter_sections(idx_new), n1)
            r2 = get_section_ranges(filter_sections(idx_old), n2)

            s1, e1 = r1[section]
            s2, e2 = r2[section]

            tnew = extract_tables_range(p1, s1, e1)
            told = extract_tables_range(p2, s2, e2)

            fnew = merge_tables_clean(tnew, p1, s1)
            fold = merge_tables_clean(told, p2, s2)

            if not fnew.empty and not fold.empty:

                all_cols = list(set(fnew.columns) & set(fold.columns))

                key_col = st.selectbox("🔑 Key column", all_cols)

                selected_cols = st.multiselect(
                    "📌 Columns to compare",
                    [c for c in all_cols if c != key_col]
                )

                if st.button("🚀 Run Deep Comparison"):
                    st.session_state["deep_result"] = deep_compare(
                        fnew, fold, key_col, selected_cols
                    )
                    st.session_state["deep_section"] = section
                    if "param_filter" in st.session_state:
                        del st.session_state["param_filter"]

                if st.session_state["deep_result"] is not None:
                    display_comparison_results(
                        st.session_state["deep_result"],
                        st.session_state["deep_section"]
                    )

    else:
        st.warning("Upload both PDFs")
