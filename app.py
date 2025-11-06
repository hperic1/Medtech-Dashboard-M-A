import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import os
import numpy as np

# Page configuration
st.set_page_config(
    page_title="MedTech M&A & Venture Dashboard",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Color palette - muted professional colors
COLORS = {
    'ma': '#4A90E2',           # Muted blue for M&A
    'venture': '#FF9E5A',      # Light orange for Venture
    'ipo': '#50C8C8',          # Teal for IPO
    'positive': '#90EE90',     # Light green for positive deltas
    'negative': '#FFB6B6',     # Light red for negative deltas
    'text': '#2C3E50',         # Dark text
    'background': '#F8F9FA'    # Light background
}

# Custom CSS for clean, professional styling
st.markdown("""
<style>
    /* Clean, minimal styling */
    .main {
        background-color: #F8F9FA;
    }
    
    /* Metric cards */
    .metric-card {
        background-color: white;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        margin-bottom: 15px;
    }
    
    .metric-title {
        font-size: 13px;
        color: #6C757D;
        margin-bottom: 8px;
        font-weight: 500;
        letter-spacing: 0.5px;
        text-transform: uppercase;
    }
    
    .metric-value {
        font-size: 32px;
        font-weight: 600;
        color: #2C3E50;
        margin-bottom: 5px;
    }
    
    .metric-delta {
        font-size: 14px;
        font-weight: 500;
    }
    
    /* Deal cards */
    .deal-card {
        background-color: white;
        padding: 15px;
        border-radius: 8px;
        border-left: 4px solid;
        margin-bottom: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    
    .deal-company {
        font-size: 16px;
        font-weight: 600;
        color: #2C3E50;
        margin-bottom: 5px;
    }
    
    .deal-value {
        font-size: 24px;
        font-weight: 700;
        margin: 8px 0;
    }
    
    .deal-description {
        font-size: 13px;
        color: #6C757D;
        line-height: 1.5;
        margin-top: 8px;
    }
    
    /* Signal cards */
    .signal-card {
        background-color: white;
        padding: 15px 20px;
        border-radius: 8px;
        border-left: 4px solid #FFD700;
        margin-bottom: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    
    .signal-title {
        font-size: 14px;
        font-weight: 600;
        color: #2C3E50;
        margin-bottom: 5px;
    }
    
    .signal-description {
        font-size: 12px;
        color: #6C757D;
        line-height: 1.4;
    }
    
    /* Section headers */
    .section-header {
        font-size: 18px;
        font-weight: 600;
        color: #2C3E50;
        margin-bottom: 15px;
        padding-bottom: 8px;
        border-bottom: 2px solid #E9ECEF;
    }
    
    /* Dataframe styling */
    .dataframe {
        font-size: 13px;
    }
    
    /* Remove extra spacing */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    
    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        padding: 10px 24px;
        font-size: 14px;
        font-weight: 500;
    }
    
    /* Tooltip */
    .tooltip {
        position: relative;
        display: inline-block;
        cursor: help;
        color: #6C757D;
        margin-left: 5px;
    }
</style>
""", unsafe_allow_html=True)

# Data loading function
@st.cache_data
def load_data():
    """Load data from Excel file"""
    try:
        possible_paths = [
            'MedTech_YTD_StandardizedHP_Categorized.xlsx',
            './MedTech_YTD_StandardizedHP_Categorized.xlsx',
            '/mnt/user-data/uploads/MedTech_YTD_StandardizedHP_Categorized.xlsx',
            os.path.join(os.path.dirname(__file__), 'MedTech_YTD_StandardizedHP_Categorized.xlsx')
        ]
        
        excel_path = None
        for path in possible_paths:
            if os.path.exists(path):
                excel_path = path
                break
        
        if excel_path is None:
            st.error("❌ Cannot find MedTech_YTD_StandardizedHP_Categorized.xlsx")
            st.info("""
            **📁 File Location Issue**
            
            The Excel file must be in the same folder as app.py.
            
            **Looking in these locations:**
            """)
            for path in possible_paths:
                exists = "✅" if os.path.exists(path) else "❌"
                st.text(f"{exists} {path}")
            
            st.markdown("---")
            st.info("""
            **🔧 Quick Fix:**
            1. Make sure both `app.py` and `MedTech_YTD_StandardizedHP_Categorized.xlsx` are in the **same folder**
            2. Run the command from that folder: `streamlit run app.py`
            3. Or drag both files to the same location
            """)
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
        # Load all three sheets
        ma_df = pd.read_excel(excel_path, sheet_name='YTD M&A Activity')
        inv_df = pd.read_excel(excel_path, sheet_name='YTD Investment Activity')
        ipo_df = pd.read_excel(excel_path, sheet_name='YTD IPO')
        
        # Clean data
        ma_df = ma_df.fillna({'Sector': 'Other', 'Deal Value': 'Undisclosed'})
        inv_df = inv_df.fillna({'Sector': 'Other', 'Amount Raised': 'Undisclosed'})
        ipo_df = ipo_df.fillna({'Amount': 'Undisclosed'})
        
        return ma_df, inv_df, ipo_df
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def parse_value(value):
    """Parse monetary value to float"""
    if pd.isna(value) or value == 'Undisclosed':
        return 0
    try:
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').strip()
        return float(value)
    except:
        return 0

def format_currency(value):
    """Format currency values for display"""
    if pd.isna(value) or value == 'Undisclosed' or value == 0:
        return 'Undisclosed'
    try:
        val = parse_value(value)
        if val >= 1_000_000_000:
            return f"${val/1_000_000_000:.1f}B"
        elif val >= 1_000_000:
            return f"${val/1_000_000:.0f}M"
        else:
            return f"${val:,.0f}"
    except:
        return str(value)

def calculate_qoq_change(df, quarter_col='Quarter', value_col=None):
    """Calculate quarter-over-quarter change"""
    try:
        # Get unique quarters and sort
        quarters = sorted(df[quarter_col].unique(), 
                         key=lambda x: int(x.replace('Q', '')))
        
        if len(quarters) < 2:
            return 0, "N/A"
        
        current_q = quarters[-1]
        prev_q = quarters[-2]
        
        if value_col:
            # Calculate based on value
            current_val = df[df[quarter_col] == current_q][value_col].apply(parse_value).sum()
            prev_val = df[df[quarter_col] == prev_q][value_col].apply(parse_value).sum()
        else:
            # Calculate based on count
            current_val = len(df[df[quarter_col] == current_q])
            prev_val = len(df[df[quarter_col] == prev_q])
        
        if prev_val == 0:
            return 0, "N/A"
        
        change = ((current_val - prev_val) / prev_val) * 100
        return change, f"{abs(change):.1f}% vs {prev_q}"
    except:
        return 0, "N/A"

def get_top_categories(df, sector_col='Sector', value_col=None, n=3):
    """Get top N categories by deal count or value"""
    try:
        if value_col:
            # Top by value
            df_temp = df.copy()
            df_temp['_value'] = df_temp[value_col].apply(parse_value)
            result = df_temp.groupby(sector_col)['_value'].sum().sort_values(ascending=False).head(n)
        else:
            # Top by count
            result = df[sector_col].value_counts().head(n)
        return result
    except:
        return pd.Series()

def create_bar_chart(data, title, color, y_title="Value"):
    """Create muted bar chart"""
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=data.index,
        y=data.values,
        marker_color=color,
        marker_line_color=color,
        marker_line_width=0,
        hovertemplate='<b>%{x}</b><br>' + y_title + ': %{y:,.0f}<extra></extra>'
    ))
    
    fig.update_layout(
        title={'text': title, 'font': {'size': 14, 'color': '#2C3E50'}},
        xaxis={'title': '', 'showgrid': False},
        yaxis={'title': y_title, 'showgrid': True, 'gridcolor': '#E9ECEF'},
        plot_bgcolor='white',
        paper_bgcolor='white',
        height=300,
        margin=dict(l=40, r=40, t=40, b=40),
        showlegend=False
    )
    
    return fig

def create_time_trend_chart(df, quarter_col, value_col, title, color):
    """Create dual-axis time trend chart"""
    try:
        # Aggregate by quarter
        quarterly = df.groupby(quarter_col).agg({
            value_col: lambda x: sum([parse_value(v) for v in x]),
            'Company': 'count'
        }).reset_index()
        quarterly.columns = [quarter_col, 'Value', 'Count']
        
        # Sort quarters
        quarter_order = ['Q1', 'Q2', 'Q3', 'Q4']
        quarterly[quarter_col] = pd.Categorical(quarterly[quarter_col], 
                                                categories=quarter_order, ordered=True)
        quarterly = quarterly.sort_values(quarter_col)
        
        fig = go.Figure()
        
        # Add value line
        fig.add_trace(go.Scatter(
            x=quarterly[quarter_col],
            y=quarterly['Value'] / 1_000_000,  # Convert to millions
            name='Deal Value',
            mode='lines+markers',
            line=dict(color=color, width=3),
            marker=dict(size=8, color=color),
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Value: $%{y:.0f}M<extra></extra>'
        ))
        
        # Add count line
        fig.add_trace(go.Scatter(
            x=quarterly[quarter_col],
            y=quarterly['Count'],
            name='Deal Count',
            mode='lines+markers',
            line=dict(color='#95A5A6', width=2, dash='dot'),
            marker=dict(size=6, color='#95A5A6'),
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Count: %{y}<extra></extra>'
        ))
        
        fig.update_layout(
            title={'text': title, 'font': {'size': 14, 'color': '#2C3E50'}},
            xaxis={'title': '', 'showgrid': False},
            yaxis={
                'title': 'Deal Value ($M)',
                'showgrid': True,
                'gridcolor': '#E9ECEF',
                'side': 'left'
            },
            yaxis2={
                'title': 'Deal Count',
                'showgrid': False,
                'side': 'right',
                'overlaying': 'y'
            },
            plot_bgcolor='white',
            paper_bgcolor='white',
            height=300,
            margin=dict(l=40, r=40, t=40, b=40),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font={'size': 11}
            ),
            hovermode='x unified'
        )
        
        return fig
    except Exception as e:
        st.error(f"Error creating chart: {str(e)}")
        return None

def generate_signal_cards(ma_df, inv_df):
    """Dynamically generate market signal cards based on data patterns"""
    signals = []
    
    try:
        # Signal 1: Surgical Robotics activity
        ma_robotics = ma_df[ma_df['Sector'].str.contains('Surgical Robotics', na=False, case=False)]
        if len(ma_robotics) > 0:
            total_value = ma_robotics['Deal Value'].apply(parse_value).sum()
            signals.append({
                'title': 'Surgical Robotics Consolidation',
                'description': f"{len(ma_robotics)} M&A deals (${total_value/1e6:.0f}M) signal continued consolidation in surgical robotics",
                'icon': '🤖'
            })
        
        # Signal 2: AI-Driven Diagnostics
        ai_diagnostics = pd.concat([
            ma_df[ma_df['Sector'].str.contains('Diagnostics', na=False, case=False)],
            inv_df[inv_df['Sector'].str.contains('Diagnostics', na=False, case=False)]
        ])
        if len(ai_diagnostics) > 5:
            signals.append({
                'title': 'AI-Driven Diagnostics Surge',
                'description': f"{len(ai_diagnostics)} deals in diagnostics space driven by AI/ML integration momentum",
                'icon': '🧠'
            })
        
        # Signal 3: Cardiology activity
        cardio_total = pd.concat([
            ma_df[ma_df['Sector'].str.contains('Cardiology', na=False, case=False)],
            inv_df[inv_df['Sector'].str.contains('Cardiology', na=False, case=False)]
        ])
        if len(cardio_total) > 8:
            signals.append({
                'title': 'Cardiovascular Innovation Wave',
                'description': f"{len(cardio_total)} deals in cardiology reflecting increased focus on heart health technologies",
                'icon': '❤️'
            })
        
        # Signal 4: Remote Patient Monitoring
        rpm = inv_df[inv_df['Sector'].str.contains('Remote Patient Monitoring', na=False, case=False)]
        if len(rpm) > 3:
            total_funding = rpm['Amount Raised'].apply(parse_value).sum()
            signals.append({
                'title': 'Remote Patient Monitoring Expansion',
                'description': f"${total_funding/1e6:.0f}M in venture funding across {len(rpm)} companies driving RPM adoption",
                'icon': '📱'
            })
        
        # Signal 5: Large deals indicator
        large_ma = ma_df[ma_df['Deal Value'].apply(lambda x: parse_value(x) > 100_000_000)]
        if len(large_ma) > 3:
            signals.append({
                'title': 'Mega-Deal Activity',
                'description': f"{len(large_ma)} deals exceeding $100M indicate strategic consolidation by major players",
                'icon': '💰'
            })
    
    except Exception as e:
        st.warning(f"Error generating signals: {str(e)}")
    
    return signals

# Main app
def main():
    st.title("🏥 MedTech M&A & Venture Dashboard")
    
    # Load data
    ma_df, inv_df, ipo_df = load_data()
    
    if ma_df.empty or inv_df.empty:
        st.error("Unable to load data. Please ensure the Excel file is available.")
        return
    
    # Sidebar navigation and filters
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("", ["Home", "Deals", "JP Morgan", "IPO"])
    
    # Global filters in sidebar
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filters")
    
    # Time period filter
    all_quarters = sorted(set(list(ma_df['Quarter'].unique()) + list(inv_df['Quarter'].unique())),
                         key=lambda x: int(x.replace('Q', '')))
    selected_quarters = st.sidebar.multiselect(
        "Quarter",
        all_quarters,
        default=all_quarters
    )
    
    # Category filter
    all_categories = sorted(set(list(ma_df['Sector'].dropna().unique()) + 
                                list(inv_df['Sector'].dropna().unique())))
    selected_categories = st.sidebar.multiselect(
        "Category",
        all_categories,
        default=all_categories
    )
    
    # Deal type filter
    deal_types = ['All', 'Merger', 'Acquisition']
    selected_deal_type = st.sidebar.selectbox("Deal Type (M&A)", deal_types)
    
    # Apply filters
    ma_filtered = ma_df.copy()
    inv_filtered = inv_df.copy()
    
    if selected_quarters:
        ma_filtered = ma_filtered[ma_filtered['Quarter'].isin(selected_quarters)]
        inv_filtered = inv_filtered[inv_filtered['Quarter'].isin(selected_quarters)]
    
    if selected_categories:
        ma_filtered = ma_filtered[ma_filtered['Sector'].isin(selected_categories)]
        inv_filtered = inv_filtered[inv_filtered['Sector'].isin(selected_categories)]
    
    if selected_deal_type != 'All':
        ma_filtered = ma_filtered[ma_filtered['Deal Type (Merger / Acquisition)'] == selected_deal_type]
    
    # Password-protected upload
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔒 Upload New Dataset")
    password = st.sidebar.text_input("Password", type="password", key="upload_password")
    
    if password == "BeaconOne":
        uploaded_file = st.sidebar.file_uploader(
            "Upload Excel file",
            type=['xlsx'],
            help="Must contain sheets: 'YTD M&A Activity', 'YTD Investment Activity', 'YTD IPO'"
        )
        
        if uploaded_file:
            try:
                # Save uploaded file
                save_path = 'MedTech_YTD_StandardizedHP_Categorized.xlsx'
                with open(save_path, 'wb') as f:
                    f.write(uploaded_file.read())
                st.sidebar.success("✅ File uploaded! Please refresh the page to see updated data.")
            except Exception as e:
                st.sidebar.error(f"Error uploading file: {str(e)}")
    elif password and password != "BeaconOne":
        st.sidebar.error("❌ Incorrect password")
    
    # Route to appropriate page
    if page == "Home":
        show_home(ma_filtered, inv_filtered)
    elif page == "Deals":
        show_deals(ma_filtered, inv_filtered)
    elif page == "JP Morgan":
        show_jp_morgan(ma_df, inv_df)
    elif page == "IPO":
        show_ipo(ipo_df)

def show_home(ma_df, inv_df):
    """Home page with split-screen M&A and Venture overview"""
    st.header("Market Overview")
    
    # Create two columns for M&A and Venture
    col_ma, col_venture = st.columns(2)
    
    # ===== LEFT COLUMN: M&A =====
    with col_ma:
        st.markdown(f"<div class='section-header' style='color: {COLORS['ma']}'>M&A Activity</div>", 
                   unsafe_allow_html=True)
        
        # KPI Cards
        total_deals_ma = len(ma_df)
        total_value_ma = ma_df['Deal Value'].apply(parse_value).sum()
        avg_deal_size_ma = total_value_ma / total_deals_ma if total_deals_ma > 0 else 0
        
        # QoQ change
        qoq_change_ma, qoq_label_ma = calculate_qoq_change(ma_df, 'Quarter')
        delta_color_ma = COLORS['positive'] if qoq_change_ma >= 0 else COLORS['negative']
        
        # Display KPIs in grid
        kpi1, kpi2 = st.columns(2)
        with kpi1:
            st.metric("Total Deals YTD", f"{total_deals_ma:,}")
        with kpi2:
            st.metric("Total Value", format_currency(total_value_ma))
        
        kpi3, kpi4 = st.columns(2)
        with kpi3:
            st.metric("Avg Deal Size", format_currency(avg_deal_size_ma))
        with kpi4:
            delta_prefix = "+" if qoq_change_ma >= 0 else ""
            st.metric("QoQ Change", 
                     f"{delta_prefix}{qoq_change_ma:.1f}%",
                     delta=qoq_label_ma)
        
        # Top 3 Categories
        st.markdown("**Top Categories by Deal Count**")
        top_cats_count = get_top_categories(ma_df, 'Sector', None, 3)
        for i, (cat, count) in enumerate(top_cats_count.items(), 1):
            st.caption(f"{i}. {cat}: {int(count)} deals")
        
        st.markdown("**Top Categories by Deal Value**")
        top_cats_value = get_top_categories(ma_df, 'Sector', 'Deal Value', 3)
        for i, (cat, val) in enumerate(top_cats_value.items(), 1):
            st.caption(f"{i}. {cat}: {format_currency(val)}")
        
        # Top 3 Deals
        st.markdown("---")
        st.markdown("**Top 3 M&A Deals**")
        
        ma_sorted = ma_df.copy()
        ma_sorted['_value_num'] = ma_sorted['Deal Value'].apply(parse_value)
        top_3_ma = ma_sorted.nlargest(3, '_value_num')
        
        for idx, row in top_3_ma.iterrows():
            deal_value = format_currency(row['Deal Value'])
            st.markdown(
                f"""<div class='deal-card' style='border-left-color: {COLORS['ma']}'>
                    <div class='deal-company'>{row['Company']}</div>
                    <div class='deal-value' style='color: {COLORS['ma']}'>{deal_value}</div>
                    <div class='deal-description'>{row['Technology/Description'][:150]}...</div>
                </div>""",
                unsafe_allow_html=True
            )
        
        # Market Overview Charts
        st.markdown("---")
        st.markdown("**Market Overview**")
        
        # Deal Value by Category
        cat_value = ma_df.groupby('Sector')['Deal Value'].apply(
            lambda x: sum([parse_value(v) for v in x]) / 1_000_000
        ).sort_values(ascending=False)
        fig1 = create_bar_chart(cat_value, "Deal Value by Category", COLORS['ma'], "Value ($M)")
        st.plotly_chart(fig1, use_container_width=True)
        
        # Deal Count by Category
        cat_count = ma_df['Sector'].value_counts()
        fig2 = create_bar_chart(cat_count, "Deal Count by Category", COLORS['ma'], "Count")
        st.plotly_chart(fig2, use_container_width=True)
        
        # Time Trends
        st.markdown("**Time Trends**")
        fig3 = create_time_trend_chart(ma_df, 'Quarter', 'Deal Value', 
                                       "M&A Activity Over Time", COLORS['ma'])
        if fig3:
            st.plotly_chart(fig3, use_container_width=True)
    
    # ===== RIGHT COLUMN: VENTURE =====
    with col_venture:
        st.markdown(f"<div class='section-header' style='color: {COLORS['venture']}'>Venture Investment</div>", 
                   unsafe_allow_html=True)
        
        # KPI Cards
        total_deals_inv = len(inv_df)
        total_value_inv = inv_df['Amount Raised'].apply(parse_value).sum()
        
        # Median round size
        inv_values = inv_df['Amount Raised'].apply(parse_value)
        median_round_size = inv_values[inv_values > 0].median() if len(inv_values[inv_values > 0]) > 0 else 0
        
        # QoQ change
        qoq_change_inv, qoq_label_inv = calculate_qoq_change(inv_df, 'Quarter')
        
        # Display KPIs in grid
        kpi1, kpi2 = st.columns(2)
        with kpi1:
            st.metric("Total Deals YTD", f"{total_deals_inv:,}")
        with kpi2:
            st.metric("Total Value", format_currency(total_value_inv))
        
        kpi3, kpi4 = st.columns(2)
        with kpi3:
            st.metric("Median Round Size", format_currency(median_round_size))
        with kpi4:
            delta_prefix = "+" if qoq_change_inv >= 0 else ""
            st.metric("QoQ Change", 
                     f"{delta_prefix}{qoq_change_inv:.1f}%",
                     delta=qoq_label_inv)
        
        # Top 3 Categories
        st.markdown("**Top Categories by Deal Count**")
        top_cats_count_inv = get_top_categories(inv_df, 'Sector', None, 3)
        for i, (cat, count) in enumerate(top_cats_count_inv.items(), 1):
            st.caption(f"{i}. {cat}: {int(count)} deals")
        
        st.markdown("**Top Categories by Deal Value**")
        top_cats_value_inv = get_top_categories(inv_df, 'Sector', 'Amount Raised', 3)
        for i, (cat, val) in enumerate(top_cats_value_inv.items(), 1):
            st.caption(f"{i}. {cat}: {format_currency(val)}")
        
        # Top 3 Deals
        st.markdown("---")
        st.markdown("**Top 3 Venture Deals**")
        
        inv_sorted = inv_df.copy()
        inv_sorted['_value_num'] = inv_sorted['Amount Raised'].apply(parse_value)
        top_3_inv = inv_sorted.nlargest(3, '_value_num')
        
        for idx, row in top_3_inv.iterrows():
            deal_value = format_currency(row['Amount Raised'])
            st.markdown(
                f"""<div class='deal-card' style='border-left-color: {COLORS['venture']}'>
                    <div class='deal-company'>{row['Company']}</div>
                    <div class='deal-value' style='color: {COLORS['venture']}'>{deal_value}</div>
                    <div class='deal-description'>{row['Technology/Description'][:150]}...</div>
                </div>""",
                unsafe_allow_html=True
            )
        
        # Market Overview Charts
        st.markdown("---")
        st.markdown("**Market Overview**")
        
        # Deal Value by Category
        cat_value_inv = inv_df.groupby('Sector')['Amount Raised'].apply(
            lambda x: sum([parse_value(v) for v in x]) / 1_000_000
        ).sort_values(ascending=False)
        fig4 = create_bar_chart(cat_value_inv, "Investment Value by Category", 
                               COLORS['venture'], "Value ($M)")
        st.plotly_chart(fig4, use_container_width=True)
        
        # Deal Count by Category
        cat_count_inv = inv_df['Sector'].value_counts()
        fig5 = create_bar_chart(cat_count_inv, "Investment Count by Category", 
                               COLORS['venture'], "Count")
        st.plotly_chart(fig5, use_container_width=True)
        
        # Time Trends
        st.markdown("**Time Trends**")
        fig6 = create_time_trend_chart(inv_df, 'Quarter', 'Amount Raised', 
                                       "Venture Activity Over Time", COLORS['venture'])
        if fig6:
            st.plotly_chart(fig6, use_container_width=True)

def show_deals(ma_df, inv_df):
    """Deals page with M&A and Venture tables"""
    st.header("Deal Tables")
    
    # Create tabs for M&A and Venture
    tab_ma, tab_venture = st.tabs(["M&A Deals", "Venture Deals"])
    
    with tab_ma:
        st.subheader("M&A Deals")
        
        # Additional filters
        col1, col2, col3 = st.columns(3)
        with col1:
            search_ma = st.text_input("🔍 Search", key='search_ma', 
                                     placeholder="Company, acquirer, technology...")
        with col2:
            sort_by_ma = st.selectbox("Sort by", 
                                     ["Deal Value (High to Low)", "Deal Value (Low to High)",
                                      "Recent First", "Company A-Z"],
                                     key='sort_ma')
        with col3:
            show_top_quartile = st.checkbox("Highlight Top Quartile", value=True)
        
        # Apply search filter
        ma_display = ma_df.copy()
        if search_ma:
            mask = ma_display.apply(
                lambda row: row.astype(str).str.contains(search_ma, case=False).any(), 
                axis=1
            )
            ma_display = ma_display[mask]
        
        # Add numeric column for sorting
        ma_display['_value_num'] = ma_display['Deal Value'].apply(parse_value)
        
        # Apply sorting
        if sort_by_ma == "Deal Value (High to Low)":
            ma_display = ma_display.sort_values('_value_num', ascending=False)
        elif sort_by_ma == "Deal Value (Low to High)":
            ma_display = ma_display.sort_values('_value_num', ascending=True)
        elif sort_by_ma == "Recent First":
            ma_display = ma_display.sort_values(['Quarter', 'Month'], ascending=False)
        else:  # Company A-Z
            ma_display = ma_display.sort_values('Company')
        
        # Highlight top quartile
        if show_top_quartile and len(ma_display) > 0:
            threshold = ma_display['_value_num'].quantile(0.75)
            
            def highlight_top(row):
                if row['_value_num'] >= threshold and row['_value_num'] > 0:
                    return ['background-color: #E8F4F8'] * len(row)
                return [''] * len(row)
        
        # Format for display
        ma_display['Deal Value'] = ma_display['Deal Value'].apply(format_currency)
        
        # Select columns to display
        display_cols = ['Company', 'Acquirer', 'Deal Type (Merger / Acquisition)', 
                       'Technology/Description', 'Deal Value', 'Sector', 'Quarter', 'Month']
        
        if show_top_quartile and len(ma_display) > 0:
            st.dataframe(
                ma_display[display_cols].style.apply(highlight_top, axis=1),
                use_container_width=True,
                height=500
            )
        else:
            st.dataframe(
                ma_display[display_cols],
                use_container_width=True,
                height=500
            )
        
        st.caption(f"Showing {len(ma_display)} of {len(ma_df)} total M&A deals")
    
    with tab_venture:
        st.subheader("Venture Investment Deals")
        
        # Additional filters
        col1, col2, col3 = st.columns(3)
        with col1:
            search_inv = st.text_input("🔍 Search", key='search_inv',
                                      placeholder="Company, investors, technology...")
        with col2:
            sort_by_inv = st.selectbox("Sort by",
                                      ["Amount (High to Low)", "Amount (Low to High)",
                                       "Recent First", "Company A-Z"],
                                      key='sort_inv')
        with col3:
            show_top_quartile_inv = st.checkbox("Highlight Top Quartile", value=True, key='top_inv')
        
        # Apply search filter
        inv_display = inv_df.copy()
        if search_inv:
            mask = inv_display.apply(
                lambda row: row.astype(str).str.contains(search_inv, case=False).any(),
                axis=1
            )
            inv_display = inv_display[mask]
        
        # Add numeric column for sorting
        inv_display['_value_num'] = inv_display['Amount Raised'].apply(parse_value)
        
        # Apply sorting
        if sort_by_inv == "Amount (High to Low)":
            inv_display = inv_display.sort_values('_value_num', ascending=False)
        elif sort_by_inv == "Amount (Low to High)":
            inv_display = inv_display.sort_values('_value_num', ascending=True)
        elif sort_by_inv == "Recent First":
            inv_display = inv_display.sort_values(['Quarter', 'Month'], ascending=False)
        else:  # Company A-Z
            inv_display = inv_display.sort_values('Company')
        
        # Highlight top quartile
        if show_top_quartile_inv and len(inv_display) > 0:
            threshold = inv_display['_value_num'].quantile(0.75)
            
            def highlight_top_inv(row):
                if row['_value_num'] >= threshold and row['_value_num'] > 0:
                    return ['background-color: #FFF4E8'] * len(row)
                return [''] * len(row)
        
        # Format for display
        inv_display['Amount Raised'] = inv_display['Amount Raised'].apply(format_currency)
        
        # Select columns to display
        display_cols = ['Company', 'Funding type (VC / PE)', 'Technology/Description',
                       'Amount Raised', 'Lead Investors', 'Sector', 'Quarter', 'Month']
        
        if show_top_quartile_inv and len(inv_display) > 0:
            st.dataframe(
                inv_display[display_cols].style.apply(highlight_top_inv, axis=1),
                use_container_width=True,
                height=500
            )
        else:
            st.dataframe(
                inv_display[display_cols],
                use_container_width=True,
                height=500
            )
        
        st.caption(f"Showing {len(inv_display)} of {len(inv_df)} total Venture deals")

def show_jp_morgan(ma_df, inv_df):
    """JP Morgan trend report page with existing style + dynamic signal cards"""
    st.header("JP Morgan MedTech Industry Report")
    
    # Keep existing JP Morgan data (hardcoded from original app)
    st.markdown("### 2025 Q1-Q3 Activity Summary")
    
    # Create 1x2 grid for charts
    col1, col2 = st.columns(2)
    
    with col1:
        # M&A Chart
        quarters = ['Q1', 'Q2', 'Q3']
        ma_values = [9200, 2100, 21700]  # JP Morgan data
        ma_counts = [57, 43, 65]
        
        fig_ma = go.Figure()
        
        fig_ma.add_trace(go.Bar(
            x=quarters,
            y=ma_values,
            name='Deal Value ($M)',
            marker_color=COLORS['ma'],
            text=[f"${v:,.0f}M" for v in ma_values],
            textposition='outside',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Deal Value: $%{y:,.0f}M<extra></extra>'
        ))
        
        fig_ma.add_trace(go.Scatter(
            x=quarters,
            y=ma_counts,
            name='Deal Count',
            mode='lines+markers+text',
            line=dict(color='#90EE90', width=3),
            marker=dict(size=10, color='#90EE90'),
            text=[str(c) for c in ma_counts],
            textposition='top center',
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Deal Count: %{y}<extra></extra>'
        ))
        
        fig_ma.update_layout(
            title='M&A Activity',
            xaxis={'title': '', 'showgrid': False},
            yaxis={
                'title': 'Deal Value ($M)',
                'side': 'left',
                'showgrid': True,
                'gridcolor': '#E9ECEF'
            },
            yaxis2={
                'title': 'Deal Count',
                'overlaying': 'y',
                'side': 'right',
                'showgrid': False
            },
            plot_bgcolor='white',
            paper_bgcolor='white',
            height=350,
            margin=dict(t=60, b=40, l=40, r=40),
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        
        st.plotly_chart(fig_ma, use_container_width=True)
    
    with col2:
        # Venture Chart
        venture_values = [3700, 2600, 2900]  # JP Morgan data
        venture_counts = [117, 90, 67]
        
        fig_venture = go.Figure()
        
        fig_venture.add_trace(go.Bar(
            x=quarters,
            y=venture_values,
            name='Deal Value ($M)',
            marker_color=COLORS['venture'],
            text=[f"${v:,.0f}M" for v in venture_values],
            textposition='outside',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Deal Value: $%{y:,.0f}M<extra></extra>'
        ))
        
        fig_venture.add_trace(go.Scatter(
            x=quarters,
            y=venture_counts,
            name='Deal Count',
            mode='lines+markers+text',
            line=dict(color='#90EE90', width=3),
            marker=dict(size=10, color='#90EE90'),
            text=[str(c) for c in venture_counts],
            textposition='top center',
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Deal Count: %{y}<extra></extra>'
        ))
        
        fig_venture.update_layout(
            title='Venture Activity',
            xaxis={'title': '', 'showgrid': False},
            yaxis={
                'title': 'Deal Value ($M)',
                'side': 'left',
                'showgrid': True,
                'gridcolor': '#E9ECEF'
            },
            yaxis2={
                'title': 'Deal Count',
                'overlaying': 'y',
                'side': 'right',
                'showgrid': False
            },
            plot_bgcolor='white',
            paper_bgcolor='white',
            height=350,
            margin=dict(t=60, b=40, l=40, r=40),
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        
        st.plotly_chart(fig_venture, use_container_width=True)
    
    # Key Market Trends
    st.markdown("---")
    st.markdown("### Key Market Trends")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**M&A Activity**")
        st.markdown("")
        st.markdown("• **Q1 2025**: 57 medtech M&A deals totaling $9.2B")
        st.markdown("• **Q2 2025**: 43 medtech M&A deals totaling $2.1B")
        st.markdown("• **Q3 2025**: 65 medtech M&A deals totaling $21.7B")
        st.markdown("")
        st.info("**Trend**: M&A activity increased through Q3 2025, surpassing full-year 2024 with strategic consolidation driving large-scale transactions")
    
    with col2:
        st.markdown("**Venture Capital**")
        st.markdown("")
        st.markdown("• **Q1 2025**: Larger rounds into fewer companies exceeded Q1 2024")
        st.markdown("• **Q2 2025**: $6.8B in H1 2025 positions sector to exceed 2024's $12.7B")
        st.markdown("• **Q3 2025**: Weaker Q2/Q3 amid challenging venture environment")
        st.markdown("")
        st.info("**Trend**: Late-stage rounds dominate at $7.9B YTD, while early-stage funding remains selective")
    
    # Dynamic Signal Cards
    st.markdown("---")
    st.markdown("### Market Signals")
    st.caption("Dynamically generated insights based on current deal data")
    
    signals = generate_signal_cards(ma_df, inv_df)
    
    # Display signal cards in grid
    cols = st.columns(2)
    for i, signal in enumerate(signals):
        with cols[i % 2]:
            st.markdown(
                f"""<div class='signal-card'>
                    <div class='signal-title'>{signal['icon']} {signal['title']}</div>
                    <div class='signal-description'>{signal['description']}</div>
                </div>""",
                unsafe_allow_html=True
            )

def show_ipo(ipo_df):
    """IPO activity page"""
    st.header("IPO Activity")
    
    if ipo_df.empty:
        st.warning("No IPO data available")
        return
    
    # KPI Cards
    col1, col2, col3, col4 = st.columns(4)
    
    total_ipos = len(ipo_df)
    total_proceeds = ipo_df['Amount'].apply(parse_value).sum()
    avg_proceeds = total_proceeds / total_ipos if total_ipos > 0 else 0
    
    with col1:
        st.metric("Total IPOs YTD", f"{total_ipos:,}")
    with col2:
        st.metric("Total Proceeds", format_currency(total_proceeds))
    with col3:
        st.metric("Avg Proceeds", format_currency(avg_proceeds))
    with col4:
        # Most recent IPO
        if 'Date' in ipo_df.columns:
            try:
                ipo_df['_date'] = pd.to_datetime(ipo_df['Date'])
                most_recent = ipo_df.nlargest(1, '_date')['Company'].values[0]
                st.metric("Most Recent", most_recent)
            except:
                st.metric("Most Recent", "N/A")
    
    st.markdown("---")
    
    # Charts in columns
    col1, col2 = st.columns(2)
    
    with col1:
        # IPO Count by Quarter
        ipo_by_quarter = ipo_df['Quarter'].value_counts().sort_index()
        
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(
            x=ipo_by_quarter.index,
            y=ipo_by_quarter.values,
            marker_color=COLORS['ipo'],
            text=ipo_by_quarter.values,
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>IPOs: %{y}<extra></extra>'
        ))
        
        fig1.update_layout(
            title='IPO Count by Quarter',
            xaxis={'title': '', 'showgrid': False},
            yaxis={'title': 'Number of IPOs', 'showgrid': True, 'gridcolor': '#E9ECEF'},
            plot_bgcolor='white',
            paper_bgcolor='white',
            height=300,
            margin=dict(l=40, r=40, t=40, b=40)
        )
        
        st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        # Proceeds by Quarter
        proceeds_by_quarter = ipo_df.groupby('Quarter')['Amount'].apply(
            lambda x: sum([parse_value(v) for v in x]) / 1_000_000
        ).sort_index()
        
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            x=proceeds_by_quarter.index,
            y=proceeds_by_quarter.values,
            marker_color=COLORS['ipo'],
            text=[f"${v:.0f}M" for v in proceeds_by_quarter.values],
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Proceeds: $%{y:.0f}M<extra></extra>'
        ))
        
        fig2.update_layout(
            title='Total Proceeds by Quarter',
            xaxis={'title': '', 'showgrid': False},
            yaxis={'title': 'Proceeds ($M)', 'showgrid': True, 'gridcolor': '#E9ECEF'},
            plot_bgcolor='white',
            paper_bgcolor='white',
            height=300,
            margin=dict(l=40, r=40, t=40, b=40)
        )
        
        st.plotly_chart(fig2, use_container_width=True)
    
    # IPO Table
    st.markdown("---")
    st.subheader("IPO Listings")
    
    # Format for display
    ipo_display = ipo_df.copy()
    ipo_display['Amount'] = ipo_display['Amount'].apply(format_currency)
    
    # Sort by date if available
    if 'Date' in ipo_display.columns:
        try:
            ipo_display['_date'] = pd.to_datetime(ipo_display['Date'])
            ipo_display = ipo_display.sort_values('_date', ascending=False)
        except:
            pass
    
    # Select columns
    display_cols = ['Company', 'Type', 'Technology', 'Investors/Deal Details', 
                   'Amount', 'Quarter']
    if 'Date' in ipo_display.columns:
        display_cols.append('Date')
    
    # Highlight largest IPOs
    if len(ipo_display) > 0:
        ipo_display['_value_num'] = ipo_df['Amount'].apply(parse_value)
        threshold = ipo_display['_value_num'].quantile(0.75)
        
        def highlight_large(row):
            if row['_value_num'] >= threshold and row['_value_num'] > 0:
                return ['background-color: #E0F5F5'] * len(row)
            return [''] * len(row)
        
        st.dataframe(
            ipo_display[display_cols].style.apply(highlight_large, axis=1),
            use_container_width=True,
            height=400
        )
    else:
        st.dataframe(ipo_display[display_cols], use_container_width=True, height=400)
    
    st.caption(f"Showing all {len(ipo_df)} IPOs")

if __name__ == "__main__":
    main()