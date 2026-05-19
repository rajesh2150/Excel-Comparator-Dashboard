"""
Employee Excel Comparator - Streamlit Application
A professional enterprise dashboard for comparing Base and SOW Excel files
with intelligent employee matching and analytics.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import tempfile
import os
from datetime import datetime
from comparator import ExcelComparator

# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Employee Excel Comparator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# CUSTOM STYLING
# =========================================================

st.markdown(
    """
    <style>
    * {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    .main {
        padding: 2rem;
    }
    
    .header-title {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f77b4;
        margin-bottom: 0.5rem;
    }
    
    .header-subtitle {
        font-size: 1.1rem;
        color: #666;
        margin-bottom: 2rem;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        text-align: center;
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        margin: 0.5rem 0;
    }
    
    .metric-label {
        font-size: 0.9rem;
        opacity: 0.9;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .perfect-match {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    }
    
    .mismatch-project {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    }
    
    .mismatch-name {
        background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
    }
    
    .updated {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
    }
    
    .section-title {
        font-size: 1.5rem;
        font-weight: 600;
        color: #333;
        margin-top: 2rem;
        margin-bottom: 1rem;
        border-bottom: 3px solid #1f77b4;
        padding-bottom: 0.5rem;
    }
    
    .status-badge-perfect {
        background-color: #d4edda;
        color: #155724;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        font-weight: 600;
    }
    
    .status-badge-mismatch-project {
        background-color: #fff3cd;
        color: #856404;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        font-weight: 600;
    }
    
    .status-badge-mismatch-name {
        background-color: #f8d7da;
        color: #721c24;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# SESSION STATE INITIALIZATION
# =========================================================


def initialize_session_state():
    """Initialize session state variables"""
    if "comparator" not in st.session_state:
        st.session_state.comparator = None
    if "comparison_done" not in st.session_state:
        st.session_state.comparison_done = False
    if "result_df" not in st.session_state:
        st.session_state.result_df = None
    if "stats" not in st.session_state:
        st.session_state.stats = None


initialize_session_state()

# =========================================================
# UTILITY FUNCTIONS
# =========================================================


def format_number(num):
    """Format number for display"""
    return f"{int(num):,}"


def get_status_badge(status):
    """Get HTML badge for status"""
    if status == "Perfect":
        return '<span class="status-badge-perfect">✓ Perfect</span>'
    elif status == "Project MisMatched":
        return '<span class="status-badge-mismatch-project">⚠ Project Mismatch</span>'
    else:
        return '<span class="status-badge-mismatch-name">✗ Name Mismatch</span>'


def create_pie_chart_results(stats):
    """Create pie chart for result distribution"""
    labels = ["Perfect Matches", "Project Mismatches", "Name Mismatches"]
    values = [
        stats["perfect_matches"],
        stats["project_mismatches"],
        stats["name_mismatches"],
    ]
    colors = ["#38ef7d", "#f5576c", "#fee140"]

    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=values,
                marker=dict(colors=colors),
                textposition="inside",
                textinfo="label+value+percent",
                hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>",
            )
        ]
    )

    fig.update_layout(
        height=400,
        font=dict(size=12),
        title="Comparison Results Distribution",
        showlegend=True,
    )

    return fig


def create_bar_chart_emp_ids(stats):
    """Create bar chart for employee ID updates"""
    labels = ["Blank IDs Found", "IDs Updated", "Existing IDs"]
    values = [
        stats["blank_emp_ids_found"],
        stats["updated_emp_ids"],
        stats["blank_emp_ids_found"] - stats["updated_emp_ids"],
    ]
    colors = ["#f093fb", "#4facfe", "#a8edea"]

    fig = go.Figure(
        data=[
            go.Bar(
                x=labels,
                y=values,
                marker=dict(color=colors),
                text=values,
                textposition="outside",
                hovertemplate="<b>%{x}</b><br>Count: %{y}<extra></extra>",
            )
        ]
    )

    fig.update_layout(
        height=400,
        font=dict(size=12),
        title="Employee ID Update Summary",
        yaxis_title="Count",
        xaxis_title="Category",
        showlegend=False,
    )

    return fig


def create_donut_chart_success(stats):
    """Create donut chart for success percentage"""
    total = stats["total_records"]
    success = stats["perfect_matches"]
    success_pct = (success / total * 100) if total > 0 else 0

    fig = go.Figure(
        data=[
            go.Pie(
                labels=["Perfect Matches", "Others"],
                values=[success, total - success],
                marker=dict(colors=["#11998e", "#e0e0e0"]),
                hole=0.4,
                textposition="inside",
                textinfo="value",
                hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>",
            )
        ]
    )

    fig.update_layout(
        height=400,
        font=dict(size=12),
        title=f"Success Rate: {success_pct:.1f}%",
        showlegend=True,
    )

    return fig


def display_metrics(stats):
    """Display summary metrics as cards"""
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Total Records</div>
                <div class="metric-value">{format_number(stats["total_records"])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col2:
        st.markdown(
            f"""
            <div class="metric-card perfect-match">
                <div class="metric-label">Perfect Matches</div>
                <div class="metric-value">{format_number(stats["perfect_matches"])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col3:
        st.markdown(
            f"""
            <div class="metric-card mismatch-project">
                <div class="metric-label">Project Mismatches</div>
                <div class="metric-value">{format_number(stats["project_mismatches"])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    col4, col5, col6 = st.columns(3)

    with col4:
        st.markdown(
            f"""
            <div class="metric-card mismatch-name">
                <div class="metric-label">Name Mismatches</div>
                <div class="metric-value">{format_number(stats["name_mismatches"])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col5:
        st.markdown(
            f"""
            <div class="metric-card updated">
                <div class="metric-label">Blank IDs Found</div>
                <div class="metric-value">{format_number(stats["blank_emp_ids_found"])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col6:
        st.markdown(
            f"""
            <div class="metric-card updated">
                <div class="metric-label">IDs Updated</div>
                <div class="metric-value">{format_number(stats["updated_emp_ids"])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# =========================================================
# SIDEBAR
# =========================================================

st.sidebar.markdown("# 📁 File Operations")

st.sidebar.markdown("---")

with st.sidebar.expander("📤 Upload Files", expanded=True):
    st.markdown("### Base Excel File")
    base_file = st.file_uploader(
        "Upload Base Abstract Excel",
        type=["xlsx", "xlsb"],
        key="base_upload",
        help="Upload the Base Abstract Excel file (.xlsx or .xlsb)",
    )

    st.markdown("### SOW Excel File")
    sow_file = st.file_uploader(
        "Upload SOW Work Orders Excel",
        type=["xlsx"],
        key="sow_upload",
        help="Upload the SOW Suppliers & Work Orders Excel file (.xlsx)",
    )

st.sidebar.markdown("---")

# Action buttons
col_run, col_download, col_report = st.sidebar.columns(3)

with col_run:
    run_comparison = st.button(
        "🔄 Run Comparison",
        use_container_width=True,
        key="run_btn",
        disabled=base_file is None or sow_file is None,
    )

with col_download:
    download_btn = st.button(
        "⬇️ Download",
        use_container_width=True,
        key="download_btn",
        disabled=not st.session_state.comparison_done,
    )

with col_report:
    export_btn = st.button(
        "📄 Report",
        use_container_width=True,
        key="export_btn",
        disabled=not st.session_state.comparison_done,
    )

st.sidebar.markdown("---")
st.sidebar.markdown("### About")
st.sidebar.info(
    """
    **Employee Excel Comparator**
    
    This tool compares employee data between Base and SOW Excel files with intelligent name matching.
    
    **Features:**
    - Reverse name matching
    - Case-insensitive comparison
    - Project ID validation
    - Automated Emp ID updates
    - Detailed analytics
    """
)

# =========================================================
# MAIN CONTENT
# =========================================================

# Header
st.markdown(
    '<div class="header-title">📊 Employee Excel Comparator Dashboard</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="header-subtitle">Compare Base and SOW Excel files with intelligent employee matching</div>',
    unsafe_allow_html=True,
)

# =========================================================
# FILE PROCESSING LOGIC
# =========================================================

if run_comparison:
    if base_file is not None and sow_file is not None:
        with st.spinner("⏳ Processing files..."):
            try:
                # Create temporary files
                with tempfile.TemporaryDirectory() as tmpdir:
                    base_path = os.path.join(tmpdir, "base_file")
                    sow_path = os.path.join(tmpdir, "sow_file.xlsx")

                    # Determine base file extension
                    if base_file.name.endswith(".xlsb"):
                        base_path += ".xlsb"
                    else:
                        base_path += ".xlsx"

                    # Save uploaded files
                    with open(base_path, "wb") as f:
                        f.write(base_file.getbuffer())

                    with open(sow_path, "wb") as f:
                        f.write(sow_file.getbuffer())

                    # Run comparison
                    comparator = ExcelComparator()
                    comparator.load_files(base_path, sow_path)
                    result_df = comparator.compare()
                    stats = comparator.get_stats()

                    # Store in session state
                    st.session_state.comparator = comparator
                    st.session_state.result_df = result_df
                    st.session_state.stats = stats
                    st.session_state.comparison_done = True

                st.success("✅ Comparison completed successfully!")

            except Exception as e:
                st.error(f"❌ Error during comparison: {str(e)}")
                st.session_state.comparison_done = False

# =========================================================
# RESULTS DISPLAY
# =========================================================

if st.session_state.comparison_done and st.session_state.stats is not None:
    stats = st.session_state.stats

    # Summary Metrics
    st.markdown(
        '<div class="section-title">📈 Summary Metrics</div>', unsafe_allow_html=True
    )
    display_metrics(stats)

    # Visualizations
    st.markdown(
        '<div class="section-title">📊 Analytics & Visualizations</div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:
        fig_pie = create_pie_chart_results(stats)
        st.plotly_chart(fig_pie, use_container_width=True)

    with col2:
        fig_bar = create_bar_chart_emp_ids(stats)
        st.plotly_chart(fig_bar, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        fig_donut = create_donut_chart_success(stats)
        st.plotly_chart(fig_donut, use_container_width=True)

    with col4:
        # Statistics summary
        st.markdown("### 📋 Detailed Statistics")
        stats_text = f"""
        - **Total Records Processed:** {format_number(stats["total_records"])}
        - **Perfect Matches:** {format_number(stats["perfect_matches"])} ({stats["perfect_matches"] / max(stats["total_records"], 1) * 100:.1f}%)
        - **Project Mismatches:** {format_number(stats["project_mismatches"])} ({stats["project_mismatches"] / max(stats["total_records"], 1) * 100:.1f}%)
        - **Name Mismatches:** {format_number(stats["name_mismatches"])} ({stats["name_mismatches"] / max(stats["total_records"], 1) * 100:.1f}%)
        - **Blank Employee IDs Found:** {format_number(stats["blank_emp_ids_found"])}
        - **Employee IDs Updated:** {format_number(stats["updated_emp_ids"])}
        """
        st.markdown(stats_text)

    # Data Table Preview
    st.markdown(
        '<div class="section-title">📋 Results Preview</div>', unsafe_allow_html=True
    )

    result_display_df = st.session_state.result_df.copy()

    # Search functionality
    col_search, col_filter = st.columns(2)

    with col_search:
        search_term = st.text_input("🔍 Search by employee name...", "")

    with col_filter:
        filter_status = st.selectbox(
            "Filter by result status...",
            ["All", "Perfect", "Project MisMatched", "Name mismatched"],
        )

    # Apply filters
    if search_term:
        result_display_df = result_display_df[
            result_display_df.iloc[:, 0].str.contains(search_term, case=False, na=False)
        ]

    if filter_status != "All":
        result_display_df = result_display_df[
            result_display_df["Result"] == filter_status
        ]

    # Pagination
    rows_per_page = st.select_slider(
        "Rows per page", options=[10, 25, 50, 100], value=25
    )
    total_pages = (len(result_display_df) // rows_per_page) + (
        1 if len(result_display_df) % rows_per_page > 0 else 0
    )

    if total_pages > 0:
        page = st.select_slider(
            "Page", options=list(range(1, total_pages + 1)), value=1
        )
        start_idx = (page - 1) * rows_per_page
        end_idx = start_idx + rows_per_page

        display_df = result_display_df.iloc[start_idx:end_idx].copy()

        st.dataframe(
            display_df,
            use_container_width=True,
            height=400,
        )

        st.caption(
            f"Showing {start_idx + 1} to {min(end_idx, len(result_display_df))} of {len(result_display_df)} records"
        )

    # Download Results
    st.markdown(
        '<div class="section-title">💾 Export Results</div>', unsafe_allow_html=True
    )

    col_xlsx, col_csv, col_summary = st.columns(3)

    with col_xlsx:
        # Generate Excel file
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            st.session_state.result_df.to_excel(
                writer, index=False, sheet_name="Results"
            )

        st.download_button(
            label="📥 Download Excel (.xlsx)",
            data=output_buffer.getvalue(),
            file_name=f"Updated_SOW_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_csv:
        # Generate CSV file
        csv_buffer = result_display_df.to_csv(index=False)

        st.download_button(
            label="📥 Download CSV",
            data=csv_buffer,
            file_name=f"Updated_SOW_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with col_summary:
        # Generate summary report
        summary_text = f"""
Comparison Report - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
{"=" * 50}

SUMMARY STATISTICS
Total Records: {format_number(stats["total_records"])}
Perfect Matches: {format_number(stats["perfect_matches"])}
Project Mismatches: {format_number(stats["project_mismatches"])}
Name Mismatches: {format_number(stats["name_mismatches"])}
Blank IDs Found: {format_number(stats["blank_emp_ids_found"])}
Updated IDs: {format_number(stats["updated_emp_ids"])}

PERCENTAGES
Perfect Match Rate: {stats["perfect_matches"] / max(stats["total_records"], 1) * 100:.2f}%
Project Mismatch Rate: {stats["project_mismatches"] / max(stats["total_records"], 1) * 100:.2f}%
Name Mismatch Rate: {stats["name_mismatches"] / max(stats["total_records"], 1) * 100:.2f}%
Update Success Rate: {stats["updated_emp_ids"] / max(stats["blank_emp_ids_found"], 1) * 100:.2f}%
"""

        st.download_button(
            label="📄 Download Report (.txt)",
            data=summary_text,
            file_name=f"Comparison_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain",
            use_container_width=True,
        )

else:
    st.info(
        """
        👋 **Welcome to Employee Excel Comparator!**
        
        To get started:
        1. Upload your Base Abstract Excel file (sidebar)
        2. Upload your SOW Suppliers & Work Orders Excel file (sidebar)
        3. Click "Run Comparison" to process the files
        
        The system will:
        - Match employees using intelligent name matching
        - Update missing Employee IDs
        - Validate Project IDs
        - Generate detailed analytics
        """
    )

# =========================================================
# FOOTER
# =========================================================

st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #666; font-size: 0.9rem; padding: 1rem 0;'>
    <p>Employee Excel Comparator | Enterprise Edition | 2026</p>
    <p>Built By Rajesh Korlapati</p>
    </div>
    """,
    unsafe_allow_html=True,
)
