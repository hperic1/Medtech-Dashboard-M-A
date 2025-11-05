import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import os
import re
import io

# Page configuration
st.set_page_config(
    page_title="MedTech M&A & Venture Dashboard",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for full-width tables
st.markdown("""
<style>
    .dataframe {
        width: 100% !important;
    }
    div[data-testid="stDataFrame"] > div {
        width: 100% !important;
    }
    .stDataFrame {
        width: 100%;
    }
    /* Make tables span full width */
    .element-container {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# Data loading function
@st.cache_data
def load_data():
    """Load data from Excel file"""
    try:
        # Try multiple possible file paths - INCLUDING data folder
        possible_paths = [
            'data/MedTech_YTD_Standardized.xlsx',
            './data/MedTech_YTD_Standardized.xlsx',
            'MedTech_YTD_Standardized.xlsx',
            '/mnt/project/MedTech_YTD_Standardized.xlsx',
            os.path.join(os.path.dirname(__file__), 'data', 'MedTech_YTD_Standardized.xlsx'),
            os.path.join(os.path.dirname(__file__), 'MedTech_YTD_Standardized.xlsx')
        ]
        
        excel_path = None
        for path in possible_paths:
            if os.path.exists(path):
                excel_path = path
                break
        
        if excel_path is None:
            st.error("❌ Cannot find MedTech_YTD_Standardized.xlsx. Please ensure the file is in the 'data' folder or same directory as app.py")
            st.info("📁 Looking in these locations:\n" + "\n".join(f"- {p}" for p in possible_paths))
            return pd.DataFrame(), pd.DataFrame()
        
        # Load M&A data
        ma_df = pd.read_excel(excel_path, sheet_name='YTD M&A Activity')
        
        # Load Investment data
        inv_df = pd.read_excel(excel_path, sheet_name='YTD Investment Activity')
        
        # Clean and standardize data
        ma_df = ma_df.fillna('Undisclosed')
        inv_df = inv_df.fillna('Undisclosed')
        
        return ma_df, inv_df
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        st.info("💡 Make sure your Excel file has sheets named 'YTD M&A Activity' and 'YTD Investment Activity'")
        return pd.DataFrame(), pd.DataFrame()

def save_data(ma_df, inv_df):
    """Save data back to Excel file with backup for undo"""
    try:
        possible_paths = [
            'data/MedTech_YTD_Standardized.xlsx',
            './data/MedTech_YTD_Standardized.xlsx',
            'MedTech_YTD_Standardized.xlsx',
            '/mnt/project/MedTech_YTD_Standardized.xlsx',
            os.path.join(os.path.dirname(__file__), 'data', 'MedTech_YTD_Standardized.xlsx'),
            os.path.join(os.path.dirname(__file__), 'MedTech_YTD_Standardized.xlsx')
        ]
        
        excel_path = None
        for path in possible_paths:
            if os.path.exists(path):
                excel_path = path
                break
        
        if excel_path is None:
            os.makedirs('data', exist_ok=True)
            excel_path = 'data/MedTech_YTD_Standardized.xlsx'
        
        # Create backup before saving
        backup_path = excel_path.replace('.xlsx', '_backup.xlsx')
        if os.path.exists(excel_path):
            import shutil
            shutil.copy2(excel_path, backup_path)
            st.session_state.last_backup_time = pd.Timestamp.now()
        
        # Save with correct sheet names
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
            ma_df.to_excel(writer, sheet_name='YTD M&A Activity', index=False)
            inv_df.to_excel(writer, sheet_name='YTD Investment Activity', index=False)
        
        st.session_state.changes_made = True
        return True
    except Exception as e:
        st.error(f"Error saving data: {str(e)}")
        return False

def undo_last_action():
    """Restore data from backup file"""
    try:
        possible_paths = [
            'data/MedTech_YTD_Standardized.xlsx',
            './data/MedTech_YTD_Standardized.xlsx',
            'MedTech_YTD_Standardized.xlsx',
        ]
        
        excel_path = None
        for path in possible_paths:
            if os.path.exists(path):
                excel_path = path
                break
        
        if excel_path is None:
            return False, "No data file found"
        
        backup_path = excel_path.replace('.xlsx', '_backup.xlsx')
        
        if not os.path.exists(backup_path):
            return False, "No backup available to restore"
        
        import shutil
        shutil.copy2(backup_path, excel_path)
        
        if 'changes_made' in st.session_state:
            del st.session_state.changes_made
        if 'last_backup_time' in st.session_state:
            del st.session_state.last_backup_time
        
        st.cache_data.clear()
        
        return True, "Successfully restored previous version"
        
    except Exception as e:
        return False, f"Error restoring backup: {str(e)}"

def format_currency(value):
    """Format currency values"""
    if pd.isna(value) or value == 'Undisclosed':
        return 'Undisclosed'
    try:
        value = float(str(value).replace('$', '').replace('B', '').replace('M', '').replace(',', ''))
        if value >= 1000:
            return f"${value/1000:.1f}B"
        elif value > 0:
            return f"${value:.0f}M"
        else:
            return 'Undisclosed'
    except:
        return str(value)

def create_quarterly_chart(df, value_col, title):
    """Create quarterly stacked bar chart with deal count overlay"""
    try:
        quarterly_data = df.groupby('Quarter').agg({
            value_col: lambda x: sum([float(str(v).replace('$', '').replace('B', '').replace('M', '').replace(',', '')) 
                                     if v != 'Undisclosed' else 0 for v in x]),
            'Company': 'count'
        }).reset_index()
        quarterly_data.columns = ['Quarter', 'Total_Value', 'Deal_Count']
        
        quarter_order = ['Q1', 'Q2', 'Q3', 'Q4']
        quarterly_data['Quarter'] = pd.Categorical(quarterly_data['Quarter'], categories=quarter_order, ordered=True)
        quarterly_data = quarterly_data.sort_values('Quarter')
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=quarterly_data['Quarter'],
            y=quarterly_data['Total_Value'],
            name='Deal Value',
            marker_color='#1f77b4',
            text=[f"${v:,.0f}" for v in quarterly_data['Total_Value']],
            textposition='outside',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Deal Value: $%{y:,.0f}<br><extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            x=quarterly_data['Quarter'],
            y=quarterly_data['Deal_Count'],
            name='Deal Count',
            mode='lines+markers+text',
            line=dict(color='#ff7f0e', width=3),
            marker=dict(size=10),
            text=quarterly_data['Deal_Count'],
            textposition='top center',
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Deal Count: %{y}<br><extra></extra>'
        ))
        
        fig.update_layout(
            title=title,
            xaxis=dict(title='Quarter'),
            yaxis=dict(
                title='Total Deal Value (USD)',
                side='left',
                showgrid=True,
                range=[0, max(quarterly_data['Total_Value']) * 1.2]
            ),
            yaxis2=dict(
                title='Number of Deals',
                overlaying='y',
                side='right',
                showgrid=False,
                range=[0, max(quarterly_data['Deal_Count']) * 1.3]
            ),
            hovermode='x unified',
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            height=500,
            margin=dict(t=100, b=50, l=50, r=50)
        )
        
        return fig
    except Exception as e:
        st.error(f"Error creating chart: {str(e)}")
        return None

def create_jp_morgan_chart_by_category(category, color):
    """Create JP Morgan chart for a specific category with deal count overlay"""
    try:
        quarters = ['Q1', 'Q2', 'Q3']
        
        data_map = {
            'M&A': {
                'values': [9200, 2100, 21700],
                'counts': [57, 43, 65]
            },
            'Venture': {
                'values': [3700, 2600, 2900],
                'counts': [117, 90, 67]
            }
        }
        
        category_data = data_map.get(category, {'values': [0, 0, 0], 'counts': [0, 0, 0]})
        values = category_data['values']
        counts = category_data['counts']
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=quarters,
            y=values,
            name='Deal Value ($M)',
            marker_color=color,
            text=[format_currency(v) for v in values],
            textposition='outside',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Deal Value: %{text}<br><extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            x=quarters,
            y=counts,
            name='Deal Count',
            mode='lines+markers+text',
            line=dict(color='#90EE90', width=3),
            marker=dict(size=10, color='#90EE90'),
            text=[str(c) if c > 0 else '' for c in counts],
            textposition='top center',
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Deal Count: %{y}<br><extra></extra>'
        ))
        
        fig.update_layout(
            title=f'{category} Activity',
            xaxis=dict(title='Quarter'),
            yaxis=dict(
                title='Deal Value (Millions USD)',
                side='left',
                showgrid=True,
                range=[0, max(values) * 1.2]
            ),
            yaxis2=dict(
                title='Number of Deals',
                overlaying='y',
                side='right',
                showgrid=False,
                range=[0, max(counts) * 1.3] if max(counts) > 0 else [0, 100]
            ),
            hovermode='x unified',
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            height=350,
            margin=dict(t=80, b=50, l=50, r=50)
        )
        
        return fig
    except Exception as e:
        st.error(f"Error creating {category} chart: {str(e)}")
        return None

# Main app
def main():
    st.title("🏥 MedTech M&A & Venture Dashboard")
    
    ma_df, inv_df = load_data()
    
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Go to", ["Deal Activity", "JP Morgan Summary", "Data Management"])
    
    if page == "Deal Activity":
        show_deal_activity(ma_df, inv_df)
    elif page == "JP Morgan Summary":
        show_jp_morgan_summary()
    elif page == "Data Management":
        show_data_management(ma_df, inv_df)

def show_deal_activity(ma_df, inv_df):
    """Display deal activity dashboard"""
    st.header("Deal Activity Dashboard")
    
    # M&A Activity Section
    st.subheader("M&A Activity")
    
    search_ma = st.text_input("🔍 Search M&A Deals", placeholder="Search by company, acquirer, technology...", key='search_ma')
    
    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        quarters_ma = ['All'] + sorted(ma_df['Quarter'].unique().tolist())
        selected_quarter_ma = st.selectbox("Filter by Quarter", quarters_ma, key='ma_quarter')
    with filter_col2:
        months_ma = ['All'] + sorted(ma_df['Month'].unique().tolist())
        selected_month_ma = st.selectbox("Filter by Month", months_ma, key='ma_month')
    
    filtered_ma = ma_df.copy()
    if selected_quarter_ma != 'All':
        filtered_ma = filtered_ma[filtered_ma['Quarter'] == selected_quarter_ma]
    if selected_month_ma != 'All':
        filtered_ma = filtered_ma[filtered_ma['Month'] == selected_month_ma]
    
    if search_ma:
        mask = filtered_ma.apply(lambda row: row.astype(str).str.contains(search_ma, case=False).any(), axis=1)
        filtered_ma = filtered_ma[mask]
    
    tab1, tab2, tab3 = st.tabs(["📊 Table", "🏆 Top Deals", "📈 Charts"])
    
    with tab1:
        ma_display = filtered_ma.copy()
        
        def parse_to_numeric(val):
            if val == 'Undisclosed' or pd.isna(val):
                return -1
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                return float(val_str)
            except:
                return -1
        
        ma_display['_Deal_Value_Numeric'] = ma_display['Deal Value'].apply(parse_to_numeric)
        ma_display = ma_display.sort_values('_Deal_Value_Numeric', ascending=False)
        
        display_cols = [col for col in ma_display.columns if not col.startswith('_')]
        
        st.dataframe(
            ma_display[display_cols], 
            use_container_width=True, 
            height=400,
            column_config={
                "Deal Value": st.column_config.TextColumn(
                    "Deal Value",
                    help="Deal value in USD",
                )
            }
        )
    
    with tab2:
        top_deals = filtered_ma.copy()
        
        def parse_deal_value(val):
            if val == 'Undisclosed' or pd.isna(val):
                return 0
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                return float(val_str)
            except:
                return 0
        
        top_deals['Deal_Value_Numeric'] = top_deals['Deal Value'].apply(parse_deal_value)
        top_deals = top_deals.nlargest(3, 'Deal_Value_Numeric')
        
        for idx, row in top_deals.iterrows():
            formatted_value = str(row['Deal Value']) if row['Deal Value'] != 'Undisclosed' else 'Undisclosed'
            
            deal_type = row['Deal Type (Merger / Acquisition)']
            verb = "merged with" if deal_type == "Merger" else "acquired"
            
            st.markdown(f"**{row['Acquirer']} {verb} {row['Company']}**")
            st.markdown(f"<h1 style='margin-top: -10px; margin-bottom: -10px; color: #1f77b4;'>{formatted_value}</h1>", unsafe_allow_html=True)
            st.markdown("---")
    
    with tab3:
        fig = create_quarterly_chart(filtered_ma, 'Deal Value', 'M&A Activity by Quarter')
        if fig:
            st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Venture Investment Activity Section
    st.subheader("Venture Investment Activity")
    
    search_inv = st.text_input("🔍 Search Investment Deals", placeholder="Search by company, investors, technology...", key='search_inv')
    
    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        quarters_inv = ['All'] + sorted(inv_df['Quarter'].unique().tolist())
        selected_quarter_inv = st.selectbox("Filter by Quarter", quarters_inv, key='inv_quarter')
    with filter_col2:
        months_inv = ['All'] + sorted(inv_df['Month'].unique().tolist())
        selected_month_inv = st.selectbox("Filter by Month", months_inv, key='inv_month')
    
    filtered_inv = inv_df.copy()
    if selected_quarter_inv != 'All':
        filtered_inv = filtered_inv[filtered_inv['Quarter'] == selected_quarter_inv]
    if selected_month_inv != 'All':
        filtered_inv = filtered_inv[filtered_inv['Month'] == selected_month_inv]
    
    if search_inv:
        mask = filtered_inv.apply(lambda row: row.astype(str).str.contains(search_inv, case=False).any(), axis=1)
        filtered_inv = filtered_inv[mask]
    
    tab1, tab2, tab3 = st.tabs(["📊 Table", "🏆 Top Deals", "📈 Charts"])
    
    with tab1:
        inv_display = filtered_inv.copy()
        
        inv_display['_Amount_Numeric'] = inv_display['Amount Raised'].apply(
            lambda x: float(x) if pd.notna(x) and x != 'Undisclosed' and str(x).replace('.','').replace('-','').isdigit() else -1
        )
        
        inv_display = inv_display.sort_values('_Amount_Numeric', ascending=False)
        
        inv_display['Amount Raised'] = inv_display['Amount Raised'].apply(
            lambda x: f"${x:,.0f}" if pd.notna(x) and x != 'Undisclosed' and str(x).replace('.','').replace('-','').isdigit() else x
        )
        
        display_cols = [col for col in inv_display.columns if not col.startswith('_')]
        
        st.dataframe(
            inv_display[display_cols],
            use_container_width=True, 
            height=400,
            column_config={
                "Amount Raised": st.column_config.TextColumn(
                    "Amount Raised",
                    help="Investment amount in USD",
                )
            }
        )
    
    with tab2:
        top_deals = filtered_inv.copy()
        
        def parse_amount_value(val):
            if val == 'Undisclosed' or pd.isna(val):
                return 0
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                return float(val_str)
            except:
                return 0
        
        top_deals['Amount_Numeric'] = top_deals['Amount Raised'].apply(parse_amount_value)
        top_deals = top_deals.nlargest(3, 'Amount_Numeric')
        
        for idx, row in top_deals.iterrows():
            amount_val = row['Amount Raised']
            if pd.notna(amount_val) and amount_val != 'Undisclosed':
                try:
                    formatted_value = f"${float(amount_val):,.0f}"
                except:
                    formatted_value = str(amount_val)
            else:
                formatted_value = "Undisclosed"
            
            st.markdown(f"**{row['Company']}**")
            st.markdown(f"<h1 style='margin-top: -10px; margin-bottom: -10px; color: #ff7f0e;'>{formatted_value}</h1>", unsafe_allow_html=True)
            st.markdown("---")
    
    with tab3:
        fig = create_quarterly_chart(filtered_inv, 'Amount Raised', 'Venture Investment by Quarter')
        if fig:
            st.plotly_chart(fig, use_container_width=True)

def show_jp_morgan_summary():
    """Display JP Morgan summary"""
    st.header("JP Morgan MedTech Industry Report")
    
    ma_df, inv_df = load_data()
    
    def calc_quarterly_stats(df, quarter, value_col):
        q_data = df[df['Quarter'] == quarter]
        
        def parse_value(val):
            if val == 'Undisclosed' or pd.isna(val):
                return 0
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                return float(val_str)
            except:
                return 0
        
        total_value = sum(q_data[value_col].apply(parse_value))
        count = len(q_data)
        
        if total_value >= 1000000000:
            formatted_value = f"${total_value/1000000000:.1f}B"
        elif total_value >= 1000000:
            formatted_value = f"${total_value/1000000:.0f}M"
        else:
            formatted_value = "$0"
            
        return count, formatted_value
    
    beacon_stats = {}
    for q in ['Q1', 'Q2', 'Q3']:
        ma_count, ma_value = calc_quarterly_stats(ma_df, q, 'Deal Value')
        inv_count, inv_value = calc_quarterly_stats(inv_df, q, 'Amount Raised')
        beacon_stats[q] = {
            'ma_count': ma_count,
            'ma_value': ma_value,
            'inv_count': inv_count,
            'inv_value': inv_value
        }
    
    st.markdown("### 2025 Q1-Q3 Activity by Category")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig_ma = create_jp_morgan_chart_by_category('M&A', '#1f77b4')
        if fig_ma:
            st.plotly_chart(fig_ma, use_container_width=True)
    
    with col2:
        fig_venture = create_jp_morgan_chart_by_category('Venture', '#ff7f0e')
        if fig_venture:
            st.plotly_chart(fig_venture, use_container_width=True)
    
    st.markdown("---")
    st.subheader("Key Market Trends")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**M&A Activity**")
        st.markdown("")
        st.markdown("• **Q1 2025**: 57 medtech M&A deals were announced, totaling $9.2 billion")
        st.markdown("")
        st.markdown("• **Q2 2025**: 43 medtech M&A deals were announced, totaling $2.1 billion")
        st.markdown("")
        st.markdown("• **Q3 2025**: 65 medtech M&A deals were announced, totaling $21.7 billion in upfront cash and equity")
        st.markdown("")
        st.markdown("**Overarching Trend**: Medtech M&A activity increased through Q3 2025, surpassing full-year 2024 numbers")
        
    with col2:
        st.markdown("**Venture Capital**")
        st.markdown("")
        st.markdown("• **Q1 2025**: Medtech venture investment activity continued to see larger rounds into fewer companies")
        st.markdown("")
        st.markdown("• **Q2 2025**: Total venture funding reaching $6.8 billion in the first half of 2025")
        st.markdown("")
        st.markdown("• **Q3 2025**: Medtech venture funding started strong but had a weaker Q2 and Q3")
        st.markdown("")
        st.markdown("**Overarching Trend**: Late-stage venture rounds dominate at $7.9B YTD")
    
    st.markdown("---")
    st.markdown("### JPMorgan vs BeaconOne Data - Quarterly Comparison")
    
    q1_col, q2_col, q3_col = st.columns(3)
    
    with q1_col:
        st.markdown("#### Q1 2025")
        st.markdown(f"""
        <div style='background-color: #4A90E2; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>57</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q1']['ma_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #357ABD; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$9.2B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q1']['ma_value']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #50C878; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>117</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q1']['inv_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #3FA35F; padding: 20px; border-radius: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$3.7B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q1']['inv_value']}</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with q2_col:
        st.markdown("#### Q2 2025")
        st.markdown(f"""
        <div style='background-color: #4A90E2; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>43</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q2']['ma_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #357ABD; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$2.1B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q2']['ma_value']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #50C878; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>90</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q2']['inv_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #3FA35F; padding: 20px; border-radius: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$2.6B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q2']['inv_value']}</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with q3_col:
        st.markdown("#### Q3 2025")
        st.markdown(f"""
        <div style='background-color: #9B59B6; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>65</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q3']['ma_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #8E44AD; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>M&A Deal Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$21.7B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q3']['ma_value']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #50C878; padding: 20px; border-radius: 10px; margin-bottom: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Count</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>67</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q3']['inv_count']}</p>
                </div>
            </div>
        </div>
        <div style='background-color: #3FA35F; padding: 20px; border-radius: 10px;'>
            <p style='color: white; margin: 0; font-size: 12px;'>Investment Value</p>
            <div style='display: flex; justify-content: space-between; align-items: baseline;'>
                <div>
                    <p style='color: white; margin: 0; font-size: 10px;'>JPMorgan</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>$2.9B</p>
                </div>
                <div style='text-align: right;'>
                    <p style='color: white; margin: 0; font-size: 10px;'>BeaconOne</p>
                    <p style='color: white; margin: 0; font-size: 32px; font-weight: bold;'>{beacon_stats['Q3']['inv_value']}</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

def show_data_management(ma_df, inv_df):
    """Data management page"""
    st.header("Data Management")
    
    # Undo button
    if 'changes_made' in st.session_state and st.session_state.changes_made:
        col1, col2, col3 = st.columns([1, 1, 4])
        with col1:
            if st.button("↩️ Undo Last Action", type="secondary", use_container_width=True):
                success, message = undo_last_action()
                if success:
                    st.success(f"✅ {message}")
                    st.rerun()
                else:
                    st.error(f"❌ {message}")
        with col2:
            if 'last_backup_time' in st.session_state:
                st.caption(f"Last change: {st.session_state.last_backup_time.strftime('%I:%M %p')}")
        
        st.markdown("---")
    
    # Create tabs
    tab1, tab2, tab3 = st.tabs(["🌐 Web Scraper", "📊 Bulk Excel Upload", "📝 Add Manual Deals"])
    
    with tab1:
        show_web_scraper(ma_df, inv_df)
    
    with tab2:
        show_bulk_excel_upload(ma_df, inv_df)
    
    with tab3:
        show_manual_deal_entry(ma_df, inv_df)

def show_manual_deal_entry(ma_df, inv_df):
    """Manual deal entry forms"""
    st.subheader("Add New Deal Manually")
    
    deal_type = st.radio("Select Deal Type", ["M&A Activity", "Venture Investment"])
    
    if deal_type == "M&A Activity":
        st.subheader("Add M&A Deal")
        
        with st.form("ma_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                company = st.text_input("Company*")
                acquirer = st.text_input("Acquirer*")
                deal_type_ma = st.selectbox("Deal Type*", ["Acquisition", "Merger"])
            
            with col2:
                technology = st.text_area("Technology/Description*")
                deal_value = st.text_input("Deal Value (e.g., 100M, 1.5B, or Undisclosed)")
            
            col3, col4 = st.columns(2)
            with col3:
                quarter = st.selectbox("Quarter*", ["Q1", "Q2", "Q3", "Q4"])
            with col4:
                month = st.selectbox("Month*", [
                    "January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"
                ])
            
            submitted = st.form_submit_button("Add M&A Deal")
            
            if submitted:
                if company and acquirer and technology:
                    def parse_deal_input(val):
                        if not val or val.lower() == 'undisclosed':
                            return 'Undisclosed'
                        val_str = val.upper().replace('$', '').replace(',', '').strip()
                        try:
                            if 'B' in val_str:
                                num = float(val_str.replace('B', ''))
                                return f"${num * 1000000000:,.0f}"
                            elif 'M' in val_str:
                                num = float(val_str.replace('M', ''))
                                return f"${num * 1000000:,.0f}"
                            else:
                                return f"${float(val_str):,.0f}"
                        except:
                            return 'Undisclosed'
                    
                    formatted_value = parse_deal_input(deal_value)
                    
                    new_deal = pd.DataFrame({
                        'Company': [company],
                        'Acquirer': [acquirer],
                        'Deal Type (Merger / Acquisition)': [deal_type_ma],
                        'Technology/Description': [technology],
                        'Deal Value': [formatted_value],
                        'Quarter': [quarter],
                        'Month': [month]
                    })
                    
                    ma_df_updated = pd.concat([ma_df, new_deal], ignore_index=True)
                    
                    if save_data(ma_df_updated, inv_df):
                        st.success("✅ M&A deal added successfully!")
                        st.balloons()
                        st.cache_data.clear()
                        st.rerun()
                else:
                    st.error("Please fill in all required fields (*)")
    
    else:
        st.subheader("Add Venture Investment Deal")
        
        with st.form("inv_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                company = st.text_input("Company*")
                funding_type = st.selectbox("Funding Type*", ["VC", "PE"])
            
            with col2:
                technology = st.text_area("Technology/Description*")
                amount = st.text_input("Amount Raised (e.g., 50M, 1.2B, or Undisclosed)")
            
            lead_investors = st.text_input("Lead Investors")
            
            col3, col4 = st.columns(2)
            with col3:
                quarter = st.selectbox("Quarter*", ["Q1", "Q2", "Q3", "Q4"])
            with col4:
                month = st.selectbox("Month*", [
                    "January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"
                ])
            
            submitted = st.form_submit_button("Add Investment Deal")
            
            if submitted:
                if company and technology:
                    def parse_amount_input(val):
                        if not val or val.lower() == 'undisclosed':
                            return 'Undisclosed'
                        val_str = val.upper().replace('$', '').replace(',', '').strip()
                        try:
                            if 'B' in val_str:
                                num = float(val_str.replace('B', ''))
                                return int(num * 1000000000)
                            elif 'M' in val_str:
                                num = float(val_str.replace('M', ''))
                                return int(num * 1000000)
                            else:
                                return int(float(val_str))
                        except:
                            return 'Undisclosed'
                    
                    formatted_amount = parse_amount_input(amount)
                    
                    new_deal = pd.DataFrame({
                        'Company': [company],
                        'Funding type (VC / PE)': [funding_type],
                        'Technology/Description': [technology],
                        'Amount Raised': [formatted_amount],
                        'Lead Investors': [lead_investors if lead_investors else 'Undisclosed'],
                        'Quarter': [quarter],
                        'Month': [month]
                    })
                    
                    inv_df_updated = pd.concat([inv_df, new_deal], ignore_index=True)
                    
                    if save_data(ma_df, inv_df_updated):
                        st.success("✅ Investment deal added successfully!")
                        st.balloons()
                        st.cache_data.clear()
                        st.rerun()
                else:
                    st.error("Please fill in all required fields (*)")

def show_web_scraper(ma_df, inv_df):
    """Enhanced web scraper with better extraction"""
    st.subheader("🌐 Web Scraper - Extract Deals from Articles")
    
    st.info("""
    **Extract medtech deals from:**
    - News articles (e.g., MedTechDive, xtalks.com)
    - Press releases
    - Industry reports
    
    **Supported formats:**
    - Direct URL scraping
    - Copy-paste article text
    """)
    
    # Option 1: URL Scraping
    st.markdown("### Option 1: Scrape from URL")
    url = st.text_input("Enter Article URL", placeholder="https://xtalks.com/medtech-mas-in-2025-a-roundup-4100/", key="scraper_url")
    scrape_button = st.button("🔍 Scrape Deals from URL", type="primary", key="scrape_url_btn")
    
    # Option 2: Text Input
    st.markdown("### Option 2: Paste Article Text")
    article_text = st.text_area(
        "Paste article content here",
        height=200,
        placeholder="Paste the full text of the article including company names, deal values, and dates...",
        key="article_text_input"
    )
    parse_text_button = st.button("🔍 Extract Deals from Text", type="primary", key="parse_text_btn")
    
    # Handle URL scraping
    if scrape_button and url:
        with st.spinner("Fetching and analyzing article..."):
            try:
                import requests
                from bs4 import BeautifulSoup
                
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                }
                
                response = requests.get(url, headers=headers, timeout=15)
                response.raise_for_status()
                
                soup = BeautifulSoup(response.content, 'html.parser')
                
                # Remove script and style elements
                for script in soup(["script", "style"]):
                    script.decompose()
                
                text = soup.get_text()
                
                # Extract deals
                extracted_deals = extract_deals_from_text(text)
                
                if extracted_deals:
                    st.session_state.scraped_deals = extracted_deals
                    st.success(f"✅ Found {len(extracted_deals)} potential deals!")
                    st.rerun()
                else:
                    st.warning("⚠️ No deals found. Try pasting the article text directly.")
                    
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 403:
                    st.error("❌ Website blocks automated scraping (403 Forbidden)")
                    st.info("💡 **Try Option 2**: Copy the article text and paste it above")
                else:
                    st.error(f"❌ Error: {str(e)}")
            except Exception as e:
                st.error(f"❌ Error scraping URL: {str(e)}")
                st.info("💡 **Try Option 2**: Copy the article text and paste it above")
    
    # Handle text parsing
    if parse_text_button and article_text:
        with st.spinner("Extracting deals from text..."):
            extracted_deals = extract_deals_from_text(article_text)
            
            if extracted_deals:
                st.session_state.scraped_deals = extracted_deals
                st.success(f"✅ Found {len(extracted_deals)} deals!")
                st.rerun()
            else:
                st.warning("⚠️ No deals found in the text.")
    
    # Display extracted deals for review and editing
    if 'scraped_deals' in st.session_state and st.session_state.scraped_deals:
        process_extracted_deals(st.session_state.scraped_deals, ma_df, inv_df)

def extract_deals_from_text(text):
    """
    Enhanced deal extraction from article text
    Handles formats like:
    - Company A—Company B\nDescription\nDate: Jan 1, 2025\nValue: $100M
    - Company raises $50M in Series B
    - Acquirer acquires Target for $1B
    """
    deals = []
    
    # Clean text
    text = re.sub(r'\s+', ' ', text)
    
    # Month mapping
    month_map = {
        'jan': 'January', 'january': 'January',
        'feb': 'February', 'february': 'February',
        'mar': 'March', 'march': 'March',
        'apr': 'April', 'april': 'April',
        'may': 'May',
        'jun': 'June', 'june': 'June',
        'jul': 'July', 'july': 'July',
        'aug': 'August', 'august': 'August',
        'sep': 'September', 'sept': 'September', 'september': 'September',
        'oct': 'October', 'october': 'October',
        'nov': 'November', 'november': 'November',
        'dec': 'December', 'december': 'December'
    }
    
    # Split by common delimiters for deals
    deal_blocks = re.split(r'\n\n+|More info|Read more', text)
    
    for block in deal_blocks:
        if len(block) < 50:  # Skip very short blocks
            continue
        
        deal_info = {}
        
        # Pattern 1: "Company A—Company B" format (em dash)
        em_dash_pattern = r'([A-Z][A-Za-z\s&,\.\']+?)—([A-Z][A-Za-z\s&,\.\']+)'
        em_dash_match = re.search(em_dash_pattern, block)
        
        if em_dash_match:
            # Could be M&A or venture - need more context
            first_name = em_dash_match.group(1).strip()
            second_name = em_dash_match.group(2).strip()
            
            # Check if it's M&A (acquire, purchase, merge keywords)
            if re.search(r'acquir|purchas|merge|bought|deal', block, re.IGNORECASE):
                deal_info['acquirer'] = second_name
                deal_info['company'] = first_name
                deal_info['type'] = 'M&A'
            else:
                # Assume venture/partnership
                deal_info['company'] = first_name
                deal_info['acquirer'] = second_name
                deal_info['type'] = 'Venture'
        
        # Pattern 2: "Company acquired by Acquirer"
        acquired_by_pattern = r'([A-Z][A-Za-z\s&,\.\']+?)\s+(?:acquired by|purchased by|bought by)\s+([A-Z][A-Za-z\s&,\.\']+)'
        acquired_by_match = re.search(acquired_by_pattern, block, re.IGNORECASE)
        
        if acquired_by_match and 'type' not in deal_info:
            deal_info['company'] = acquired_by_match.group(1).strip()
            deal_info['acquirer'] = acquired_by_match.group(2).strip()
            deal_info['type'] = 'M&A'
        
        # Pattern 3: "Acquirer acquires Company"
        acquires_pattern = r'([A-Z][A-Za-z\s&,\.\']+?)\s+(?:acquires|acquired|purchases|purchased|bought)\s+([A-Z][A-Za-z\s&,\.\']+?)(?:\s+for|,|\.|\n)'
        acquires_match = re.search(acquires_pattern, block, re.IGNORECASE)
        
        if acquires_match and 'type' not in deal_info:
            deal_info['acquirer'] = acquires_match.group(1).strip()
            deal_info['company'] = acquires_match.group(2).strip()
            deal_info['type'] = 'M&A'
        
        # Pattern 4: "Company raises $X" (Venture)
        raises_pattern = r'([A-Z][A-Za-z\s&,\.\']+?)\s+(?:raises|raised|secures|secured)\s+(?:\$|£|€)([\d,\.]+)\s*(million|billion|M|B)?'
        raises_match = re.search(raises_pattern, block, re.IGNORECASE)
        
        if raises_match and 'type' not in deal_info:
            deal_info['company'] = raises_match.group(1).strip()
            deal_info['acquirer'] = ''
            deal_info['type'] = 'Venture'
            
            # Extract value
            amount = raises_match.group(2).replace(',', '')
            unit = raises_match.group(3) if raises_match.group(3) else 'M'
            deal_info['value'] = f"{amount}{unit}"
        
        # Extract deal value if not already found
        if 'value' not in deal_info:
            value_patterns = [
                r'Value:\s*(?:Not disclosed|Undisclosed|\$?([\d,\.]+)\s*(million|billion|M|B)?)',
                r'(?:for|valued at|worth)\s+\$?([\d,\.]+)\s*(million|billion|M|B)?',
                r'\$\s?([\d,\.]+)\s*(million|billion|M|B)',
            ]
            
            for pattern in value_patterns:
                value_match = re.search(pattern, block, re.IGNORECASE)
                if value_match:
                    if 'not disclosed' in value_match.group(0).lower() or 'undisclosed' in value_match.group(0).lower():
                        deal_info['value'] = 'Undisclosed'
                    else:
                        amount = value_match.group(1).replace(',', '') if value_match.lastindex >= 1 else '0'
                        unit = value_match.group(2) if value_match.lastindex >= 2 else 'M'
                        deal_info['value'] = f"{amount}{unit}"
                    break
        
        if 'value' not in deal_info:
            deal_info['value'] = 'Undisclosed'
        
        # Extract date
        date_pattern = r'(?:Date of Announcement|Date|Announced):\s*([A-Za-z]+\.?)\s+(\d{1,2}),?\s+(\d{4})'
        date_match = re.search(date_pattern, block, re.IGNORECASE)
        
        if date_match:
            month_str = date_match.group(1).lower().replace('.', '')
            month = month_map.get(month_str, 'January')
            year = date_match.group(3)
            
            deal_info['month'] = month
            
            # Determine quarter
            month_num = list(month_map.values()).index(month) + 1 if month in month_map.values() else 1
            deal_info['quarter'] = f"Q{(month_num - 1) // 3 + 1}"
        else:
            deal_info['month'] = 'January'
            deal_info['quarter'] = 'Q1'
        
        # Extract description (first 200 chars of main text)
        description = re.sub(r'(Date of Announcement|Date|Value|More info):.*', '', block, flags=re.IGNORECASE)
        description = description.strip()[:200]
        deal_info['description'] = description
        
        # Only add if we found at least company name
        if 'company' in deal_info and deal_info['company']:
            deals.append(deal_info)
    
    return deals

def process_extracted_deals(extracted_deals, ma_df, inv_df):
    """Display and allow editing of extracted deals before adding"""
    st.markdown("---")
    st.subheader("Review and Edit Extracted Deals")
    st.markdown(f"**{len(extracted_deals)} deals found** - Review and edit before adding to dashboard")
    
    deals_to_add = []
    deals_to_remove = []
    
    for idx, deal in enumerate(extracted_deals):
        # Determine initial deal type
        initial_type_idx = 0 if deal.get('type') == 'M&A' else 1
        
        with st.expander(f"Deal {idx + 1}: {deal.get('company', 'Unknown')}", expanded=True):
            # Header with remove button
            col_header1, col_header2 = st.columns([5, 1])
            with col_header1:
                st.markdown(f"**Deal Type:** {deal.get('type', 'Unknown')}")
            with col_header2:
                if st.button(f"🗑️ Remove", key=f"delete_{idx}", type="secondary", use_container_width=True):
                    deals_to_remove.append(idx)
            
            st.markdown("---")
            
            # Deal Type Selection - This determines which fields are shown
            deal_type_select = st.radio(
                "Select Deal Type*",
                ["M&A Activity", "Venture Investment"],
                index=initial_type_idx,
                key=f"type_{idx}",
                horizontal=True,
                help="Choose whether this is an M&A transaction or venture investment"
            )
            
            st.markdown("###")  # Add spacing
            
            # Layout fields in 2 columns based on deal type
            col1, col2 = st.columns(2)
            
            with col1:
                company = st.text_input(
                    "Company*",
                    value=deal.get('company', ''),
                    key=f"company_{idx}",
                    help="Name of the company"
                )
                
                # M&A SPECIFIC FIELDS
                if deal_type_select == "M&A Activity":
                    acquirer = st.text_input(
                        "Acquirer*",
                        value=deal.get('acquirer', ''),
                        key=f"acquirer_{idx}",
                        help="Company acquiring the target"
                    )
                    deal_subtype = st.selectbox(
                        "Deal Type*",
                        ["Acquisition", "Merger"],
                        key=f"subtype_{idx}",
                        help="Type of M&A transaction"
                    )
                
                # VENTURE SPECIFIC FIELDS
                else:
                    funding_type = st.selectbox(
                        "Funding Type*",
                        ["VC", "PE"],
                        key=f"funding_{idx}",
                        help="Venture Capital or Private Equity"
                    )
            
            with col2:
                technology = st.text_area(
                    "Technology/Description*",
                    value=deal.get('description', ''),
                    height=100,
                    key=f"tech_{idx}",
                    help="Brief description of the technology or business"
                )
                
                # Different labels and placeholders for M&A vs Venture
                if deal_type_select == "M&A Activity":
                    deal_value = st.text_input(
                        "Deal Value (e.g., 100M, 1.5B, or Undisclosed)",
                        value=deal.get('value', 'Undisclosed'),
                        key=f"value_{idx}",
                        placeholder="e.g., 100M, 1.5B, or Undisclosed",
                        help="Transaction value"
                    )
                else:
                    deal_value = st.text_input(
                        "Amount Raised (e.g., 50M, 1.2B, or Undisclosed)",
                        value=deal.get('value', 'Undisclosed'),
                        key=f"value_{idx}",
                        placeholder="e.g., 50M, 1.2B, or Undisclosed",
                        help="Amount of investment raised"
                    )
            
            # Lead Investors field (ONLY for Venture, full width)
            if deal_type_select == "Venture Investment":
                lead_investors = st.text_input(
                    "Lead Investors",
                    value=deal.get('acquirer', ''),
                    key=f"investors_{idx}",
                    help="Primary investors in the funding round (optional)"
                )
            
            # Date fields
            col3, col4 = st.columns(2)
            with col3:
                quarter = st.selectbox(
                    "Quarter*",
                    ["Q1", "Q2", "Q3", "Q4"],
                    index=["Q1", "Q2", "Q3", "Q4"].index(deal.get('quarter', 'Q1')),
                    key=f"quarter_{idx}"
                )
            with col4:
                months = ["January", "February", "March", "April", "May", "June",
                         "July", "August", "September", "October", "November", "December"]
                default_month_idx = months.index(deal.get('month', 'January'))
                month = st.selectbox(
                    "Month*",
                    months,
                    index=default_month_idx,
                    key=f"month_{idx}"
                )
            
            # Store edited deal with appropriate fields
            if deal_type_select == "M&A Activity":
                deals_to_add.append({
                    'type': 'M&A',
                    'company': company,
                    'acquirer': acquirer,
                    'deal_subtype': deal_subtype,
                    'technology': technology,
                    'value': deal_value,
                    'quarter': quarter,
                    'month': month
                })
            else:
                deals_to_add.append({
                    'type': 'Venture',
                    'company': company,
                    'funding_type': funding_type,
                    'lead_investors': lead_investors if deal_type_select == "Venture Investment" else 'Undisclosed',
                    'technology': technology,
                    'value': deal_value,
                    'quarter': quarter,
                    'month': month
                })
    
    # Remove marked deals
    for idx in reversed(deals_to_remove):
        st.session_state.scraped_deals.pop(idx)
    
    if deals_to_remove:
        st.rerun()
    
    # Action buttons
    st.markdown("---")
    col1, col2, col3 = st.columns([2, 2, 2])
    
    with col1:
        add_all_clicked = st.button("✅ Add All Deals to Dashboard", type="primary", use_container_width=True)
    
    with col2:
        if st.button("🗑️ Clear All Deals", type="secondary", use_container_width=True):
            st.session_state.scraped_deals = []
            st.success("All deals cleared!")
            st.rerun()
    
    with col3:
        st.metric("Deals to Add", len(deals_to_add))
    
    # Add deals to database
    if add_all_clicked:
        added_ma, added_inv, skipped = add_deals_to_database(deals_to_add, ma_df, inv_df)
        
        if save_data(ma_df, inv_df):
            success_msg = f"✅ Successfully added {added_ma} M&A deals and {added_inv} Venture deals!"
            if skipped:
                success_msg += f"\n\n⚠️ Skipped {len(skipped)} duplicate(s):\n" + "\n".join(f"• {s}" for s in skipped)
            
            st.success(success_msg)
            if added_ma > 0 or added_inv > 0:
                st.balloons()
            
            st.session_state.scraped_deals = []
            st.cache_data.clear()
            st.rerun()

def show_bulk_excel_upload(ma_df, inv_df):
    """Bulk CSV/Excel upload with automatic column detection and direct import"""
    st.subheader("📊 Bulk Excel/CSV Upload")
    
    st.info("""
    **Upload a CSV or Excel file with deal data**
    
    The system will automatically detect columns or let you map them:
    - **Company**, **Technology**, **Type**, **Amount**, **Date**
    - **Investors/Deal Details** (for acquirer or lead investors)
    - Supports multiple currencies (£, €, $, AU$, CA$, CHF)
    - Auto-converts to USD
    - Checks for duplicates
    - Undo available after import
    """)
    
    uploaded_file = st.file_uploader(
        "Choose CSV or Excel file",
        type=['csv', 'xlsx', 'xls'],
        key="bulk_upload",
        help="Upload a CSV or Excel file with your deal data"
    )
    
    if uploaded_file is not None:
        try:
            # Read file
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ File uploaded: {uploaded_file.name} ({len(df)} rows)")
            
            # Show preview
            st.markdown("### 📋 File Preview")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Try auto-detection first
            auto_detected = auto_detect_columns(df)
            
            if auto_detected['confidence'] == 'high':
                st.success("🎯 **Columns auto-detected!** Review mapping below:")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Detected Mappings:**")
                    for field, col in auto_detected['mapping'].items():
                        st.markdown(f"✅ **{field}**: `{col}`")
                
                with col2:
                    st.markdown("**Import Settings:**")
                    deal_type_detected = auto_detected.get('deal_type', 'Mixed')
                    st.markdown(f"📊 **Deal Type**: {deal_type_detected}")
                    st.markdown(f"💱 **Currency Conversion**: Enabled")
                    st.markdown(f"🔍 **Duplicate Check**: Enabled")
                
                # Quick import button
                col_import, col_manual = st.columns([2, 2])
                with col_import:
                    if st.button("⚡ Import All Deals Now", type="primary", use_container_width=True):
                        with st.spinner("Processing and importing deals..."):
                            imported_ma, imported_inv, skipped = direct_import_deals(
                                df, auto_detected['mapping'], ma_df, inv_df
                            )
                            
                            if save_data(ma_df, inv_df):
                                success_msg = f"✅ Successfully imported {imported_ma} M&A deals and {imported_inv} Venture deals!"
                                if skipped:
                                    success_msg += f"\n\n⚠️ Skipped {len(skipped)} duplicate(s)"
                                
                                st.success(success_msg)
                                st.balloons()
                                st.cache_data.clear()
                                st.rerun()
                
                with col_manual:
                    show_manual_mapping = st.checkbox("🔧 Manual Column Mapping", value=False)
            else:
                st.warning("⚠️ Could not auto-detect columns. Please map manually:")
                show_manual_mapping = True
            
            # Manual column mapping (if needed)
            if 'show_manual_mapping' in locals() and show_manual_mapping:
                st.markdown("---")
                st.markdown("### 🔗 Manual Column Mapping")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**Required Fields**")
                    
                    company_col = st.selectbox("Company Column*", [''] + list(df.columns), key="company_col")
                    tech_col = st.selectbox("Technology Column*", [''] + list(df.columns), key="tech_col")
                    value_col = st.selectbox("Amount Column*", [''] + list(df.columns), key="value_col")
                    date_col = st.selectbox("Date Column*", [''] + list(df.columns), key="date_col")
                
                with col2:
                    st.markdown("**Deal Type Configuration**")
                    
                    deal_type_option = st.radio(
                        "Primary Deal Type",
                        ["M&A Activity", "Venture Investment", "Mixed (has Type column)"],
                        help="Select the primary type of deals in your file"
                    )
                    
                    if deal_type_option == "Mixed (has Type column)":
                        type_col = st.selectbox("Type Column*", [''] + list(df.columns), key="type_col")
                    else:
                        type_col = None
                    
                    investors_col = st.selectbox(
                        "Investors/Acquirer Column",
                        [''] + list(df.columns),
                        key="investors_col"
                    )
                
                # Validation
                required_filled = all([company_col, tech_col, value_col, date_col])
                
                if deal_type_option == "Mixed (has Type column)":
                    required_filled = required_filled and type_col
                
                # Process button
                if st.button("🔄 Import with Manual Mapping", type="primary", disabled=not required_filled, use_container_width=True):
                    with st.spinner("Processing and importing deals..."):
                        manual_mapping = {
                            'company': company_col,
                            'technology': tech_col,
                            'amount': value_col,
                            'date': date_col,
                            'type': type_col if type_col else None,
                            'investors': investors_col if investors_col else None
                        }
                        
                        imported_ma, imported_inv, skipped = direct_import_deals(
                            df, manual_mapping, ma_df, inv_df
                        )
                        
                        if save_data(ma_df, inv_df):
                            success_msg = f"✅ Successfully imported {imported_ma} M&A deals and {imported_inv} Venture deals!"
                            if skipped:
                                success_msg += f"\n\n⚠️ Skipped {len(skipped)} duplicate(s):\n" + "\n".join(f"• {s}" for s in skipped[:5])
                            
                            st.success(success_msg)
                            st.balloons()
                            st.cache_data.clear()
                            st.rerun()
        
        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")
            st.info("💡 Make sure your file is a valid CSV or Excel file")

def auto_detect_columns(df):
    """Automatically detect column mappings based on common patterns"""
    columns_lower = {col.lower(): col for col in df.columns}
    
    mapping = {}
    confidence = 'high'
    
    # Detect Company column
    for pattern in ['company', 'firm', 'startup', 'business']:
        if pattern in columns_lower:
            mapping['company'] = columns_lower[pattern]
            break
    
    # Detect Technology column
    for pattern in ['technology', 'tech', 'description', 'desc', 'product']:
        if pattern in columns_lower:
            mapping['technology'] = columns_lower[pattern]
            break
    
    # Detect Type column
    for pattern in ['type', 'deal type', 'category', 'class']:
        if pattern in columns_lower:
            mapping['type'] = columns_lower[pattern]
            break
    
    # Detect Amount column
    for pattern in ['amount', 'value', 'price', 'deal value', 'funding']:
        if pattern in columns_lower:
            mapping['amount'] = columns_lower[pattern]
            break
    
    # Detect Date column
    for pattern in ['date', 'announcement', 'announced', 'closed']:
        if pattern in columns_lower:
            mapping['date'] = columns_lower[pattern]
            break
    
    # Detect Investors/Acquirer column
    for pattern in ['investor', 'acquirer', 'buyer', 'deal details', 'investors/deal']:
        if pattern in columns_lower:
            mapping['investors'] = columns_lower[pattern]
            break
    
    # Check confidence
    required_fields = ['company', 'technology', 'amount', 'date']
    if all(field in mapping for field in required_fields):
        confidence = 'high'
    elif len(mapping) >= 3:
        confidence = 'medium'
    else:
        confidence = 'low'
    
    # Detect deal type from Type column if present
    deal_type = 'Mixed'
    if 'type' in mapping and mapping['type'] in df.columns:
        type_values = df[mapping['type']].str.lower().unique()
        has_ma = any('merger' in str(v) or 'acquisition' in str(v) or 'grant' in str(v) 
                     or 'strategic' in str(v) or 'other' in str(v) for v in type_values)
        has_venture = any('series' in str(v) or 'seed' in str(v) or 'ipo' in str(v) 
                         or 'debt' in str(v) for v in type_values)
        
        if has_ma and has_venture:
            deal_type = 'Mixed'
        elif has_ma:
            deal_type = 'M&A Activity'
        elif has_venture:
            deal_type = 'Venture Investment'
    
    return {
        'mapping': mapping,
        'confidence': confidence,
        'deal_type': deal_type
    }

def convert_currency_to_usd(amount_str, default_currency='$'):
    """Convert various currencies to USD with rates"""
    # Currency conversion rates (approximate)
    conversion_rates = {
        '$': 1.0,        # USD
        '£': 1.27,       # GBP
        '€': 1.09,       # EUR
        'A$': 0.65,      # AUD
        'AU$': 0.65,     # AUD
        'CA$': 0.72,     # CAD
        'C$': 0.72,      # CAD
        'CHF': 1.13,     # Swiss Franc
        'NZ$': 0.60,     # New Zealand Dollar
        '¥': 0.0068,     # Japanese Yen
        '₹': 0.012,      # Indian Rupee
    }
    
    if not amount_str or pd.isna(amount_str) or str(amount_str).lower() == 'undisclosed':
        return 'Undisclosed', None
    
    amount_str = str(amount_str).strip()
    
    # Detect currency symbol
    currency = default_currency
    for symbol in conversion_rates.keys():
        if symbol in amount_str:
            currency = symbol
            break
    
    # Extract numeric value
    # Remove currency symbols and clean
    clean_str = amount_str
    for symbol in conversion_rates.keys():
        clean_str = clean_str.replace(symbol, '')
    
    clean_str = clean_str.replace(',', '').strip()
    
    try:
        # Check for million/billion indicators
        multiplier = 1
        if 'billion' in clean_str.lower() or 'b' in clean_str.lower():
            multiplier = 1000000000
            clean_str = re.sub(r'[bB]illion|[bB]', '', clean_str).strip()
        elif 'million' in clean_str.lower() or 'm' in clean_str.lower():
            multiplier = 1000000
            clean_str = re.sub(r'[mM]illion|[mM]', '', clean_str).strip()
        elif 'k' in clean_str.lower():
            multiplier = 1000
            clean_str = re.sub(r'[kK]', '', clean_str).strip()
        
        # Extract just the number
        number_match = re.search(r'[\d.]+', clean_str)
        if number_match:
            amount = float(number_match.group()) * multiplier
            
            # Convert to USD
            usd_amount = amount * conversion_rates.get(currency, 1.0)
            
            # Create conversion note if not USD
            conversion_note = None
            if currency != '$':
                conversion_note = f"Converted from {currency}{amount/multiplier:,.0f}{('B' if multiplier >= 1000000000 else 'M' if multiplier >= 1000000 else '')} to USD"
            
            return f"${usd_amount:,.0f}", conversion_note
        else:
            return 'Undisclosed', None
    except:
        return 'Undisclosed', None

def direct_import_deals(df, mapping, ma_df, inv_df):
    """Import deals directly without review step, with currency conversion and duplicate checking"""
    added_ma = 0
    added_inv = 0
    skipped_duplicates = []
    
    # Month mapping
    month_map = {
        1: 'January', 2: 'February', 3: 'March', 4: 'April',
        5: 'May', 6: 'June', 7: 'July', 8: 'August',
        9: 'September', 10: 'October', 11: 'November', 12: 'December'
    }
    
    for idx, row in df.iterrows():
        try:
            # Extract company
            company = str(row[mapping['company']]).strip()
            if not company or company == 'nan':
                continue
            
            # Extract technology
            technology = str(row[mapping['technology']]).strip()[:200] if 'technology' in mapping else ''
            
            # Extract and convert amount
            amount_raw = str(row[mapping['amount']]).strip() if 'amount' in mapping else 'Undisclosed'
            amount_usd, conversion_note = convert_currency_to_usd(amount_raw)
            
            # Determine deal type
            deal_type = 'M&A'  # Default
            if 'type' in mapping and mapping['type']:
                type_value = str(row[mapping['type']]).lower()
                if any(term in type_value for term in ['series', 'seed', 'ipo', 'debt', 'later stage']):
                    deal_type = 'Venture'
                elif any(term in type_value for term in ['grant', 'strategic', 'other']):
                    deal_type = 'M&A'
            
            # Extract investors/acquirer
            investors = 'Undisclosed'
            if 'investors' in mapping and mapping['investors']:
                investors = str(row[mapping['investors']]).strip()
            
            # Extract and parse date
            date_raw = str(row[mapping['date']]).strip() if 'date' in mapping else ''
            
            try:
                date_obj = pd.to_datetime(date_raw)
                month_num = date_obj.month
                month = month_map.get(month_num, 'January')
                quarter = f"Q{(month_num - 1) // 3 + 1}"
            except:
                month = 'January'
                quarter = 'Q1'
            
            # Check for duplicates and add to appropriate dataframe
            if deal_type == 'M&A':
                # Check duplicate
                is_duplicate = False
                for _, existing_row in ma_df.iterrows():
                    existing_company = str(existing_row['Company']).strip().lower()
                    new_company = company.lower()
                    
                    if existing_company == new_company or existing_company in new_company or new_company in existing_company:
                        if str(existing_row['Deal Value']).strip() == amount_usd:
                            is_duplicate = True
                            skipped_duplicates.append(f"M&A: {company} ({amount_usd})")
                            break
                
                if not is_duplicate:
                    # Parse acquirer from investors field
                    acquirer = investors if investors != 'Undisclosed' else 'Undisclosed'
                    
                    new_deal = pd.DataFrame({
                        'Company': [company],
                        'Acquirer': [acquirer],
                        'Deal Type (Merger / Acquisition)': ['Acquisition'],
                        'Technology/Description': [technology],
                        'Deal Value': [amount_usd],
                        'Quarter': [quarter],
                        'Month': [month]
                    })
                    ma_df = pd.concat([ma_df, new_deal], ignore_index=True)
                    added_ma += 1
            
            else:  # Venture
                # Convert amount to numeric for venture
                if amount_usd != 'Undisclosed':
                    try:
                        amount_numeric = int(float(amount_usd.replace('$', '').replace(',', '')))
                    except:
                        amount_numeric = 'Undisclosed'
                else:
                    amount_numeric = 'Undisclosed'
                
                # Check duplicate
                is_duplicate = False
                for _, existing_row in inv_df.iterrows():
                    existing_company = str(existing_row['Company']).strip().lower()
                    new_company = company.lower()
                    
                    if existing_company == new_company or existing_company in new_company or new_company in existing_company:
                        if str(existing_row['Amount Raised']).strip() == str(amount_numeric):
                            is_duplicate = True
                            display_val = f"${amount_numeric:,}" if amount_numeric != 'Undisclosed' else 'Undisclosed'
                            skipped_duplicates.append(f"Venture: {company} ({display_val})")
                            break
                
                if not is_duplicate:
                    new_deal = pd.DataFrame({
                        'Company': [company],
                        'Funding type (VC / PE)': ['VC'],
                        'Technology/Description': [technology],
                        'Amount Raised': [amount_numeric],
                        'Lead Investors': [investors],
                        'Quarter': [quarter],
                        'Month': [month]
                    })
                    inv_df = pd.concat([inv_df, new_deal], ignore_index=True)
                    added_inv += 1
        
        except Exception as e:
            st.warning(f"⚠️ Skipped row {idx + 1}: {str(e)}")
            continue
    
    return added_ma, added_inv, skipped_duplicates

def add_deals_to_database(deals_to_add, ma_df, inv_df):
    """Add processed deals to database with duplicate detection"""
    added_ma = 0
    added_inv = 0
    skipped_duplicates = []
    
    for deal in deals_to_add:
        if deal['type'] == 'M&A':
            # Parse deal value
            def parse_deal_input(val):
                if not val or val.lower() == 'undisclosed':
                    return 'Undisclosed'
                val_str = val.upper().replace('$', '').replace(',', '').strip()
                try:
                    if 'B' in val_str:
                        num = float(val_str.replace('B', '').replace('ILLION', ''))
                        return f"${num * 1000000000:,.0f}"
                    elif 'M' in val_str:
                        num = float(val_str.replace('M', '').replace('ILLION', ''))
                        return f"${num * 1000000:,.0f}"
                    else:
                        return f"${float(val_str):,.0f}"
                except:
                    return 'Undisclosed'
            
            formatted_value = parse_deal_input(deal['value'])
            
            # Check for duplicates
            is_duplicate = False
            for idx, existing_row in ma_df.iterrows():
                existing_company = str(existing_row['Company']).strip().lower()
                new_company = deal['company'].strip().lower()
                
                if existing_company == new_company or existing_company in new_company or new_company in existing_company:
                    if str(existing_row['Deal Value']).strip() == formatted_value:
                        is_duplicate = True
                        skipped_duplicates.append(f"M&A: {deal['company']} ({formatted_value})")
                        break
            
            if not is_duplicate:
                new_deal = pd.DataFrame({
                    'Company': [deal['company']],
                    'Acquirer': [deal['acquirer']],
                    'Deal Type (Merger / Acquisition)': [deal['deal_subtype']],
                    'Technology/Description': [deal['technology']],
                    'Deal Value': [formatted_value],
                    'Quarter': [deal['quarter']],
                    'Month': [deal['month']]
                })
                ma_df = pd.concat([ma_df, new_deal], ignore_index=True)
                added_ma += 1
        
        else:  # Venture
            # Parse amount
            def parse_amount_input(val):
                if not val or val.lower() == 'undisclosed':
                    return 'Undisclosed'
                val_str = val.upper().replace('$', '').replace(',', '').strip()
                try:
                    if 'B' in val_str:
                        num = float(val_str.replace('B', '').replace('ILLION', ''))
                        return int(num * 1000000000)
                    elif 'M' in val_str:
                        num = float(val_str.replace('M', '').replace('ILLION', ''))
                        return int(num * 1000000)
                    else:
                        return int(float(val_str))
                except:
                    return 'Undisclosed'
            
            formatted_amount = parse_amount_input(deal['value'])
            
            # Check for duplicates
            is_duplicate = False
            for idx, existing_row in inv_df.iterrows():
                existing_company = str(existing_row['Company']).strip().lower()
                new_company = deal['company'].strip().lower()
                
                if existing_company == new_company or existing_company in new_company or new_company in existing_company:
                    if str(existing_row['Amount Raised']).strip() == str(formatted_amount):
                        is_duplicate = True
                        display_val = f"${formatted_amount:,}" if formatted_amount != 'Undisclosed' else 'Undisclosed'
                        skipped_duplicates.append(f"Venture: {deal['company']} ({display_val})")
                        break
            
            if not is_duplicate:
                new_deal = pd.DataFrame({
                    'Company': [deal['company']],
                    'Funding type (VC / PE)': [deal['funding_type']],
                    'Technology/Description': [deal['technology']],
                    'Amount Raised': [formatted_amount],
                    'Lead Investors': [deal.get('lead_investors', 'Undisclosed')],
                    'Quarter': [deal['quarter']],
                    'Month': [deal['month']]
                })
                inv_df = pd.concat([inv_df, new_deal], ignore_index=True)
                added_inv += 1
    
    return added_ma, added_inv, skipped_duplicates

if __name__ == "__main__":
    main()