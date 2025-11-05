import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import os
import json

# Page configuration
st.set_page_config(
    page_title="MedTech M&A & Investment Activity Dashboard",
    page_icon="🥼",
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
            'data/MedTech_YTD_Standardized.xlsx',  # In data folder
            './data/MedTech_YTD_Standardized.xlsx',  # In data folder (explicit)
            'MedTech_YTD_Standardized.xlsx',  # Same directory as app.py
            'MedTech_MA_Masterlist.xlsx',  # Old filename
            './MedTech_MA_Masterlist.xlsx',
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
            st.info("🔍 Looking in these locations:\n" + "\n".join(f"- {p}" for p in possible_paths))
            return pd.DataFrame(), pd.DataFrame()
        
        # Load M&A data - NOTE: Sheet name has SPACES not underscores
        ma_df = pd.read_excel(excel_path, sheet_name='YTD M&A Activity')
        
        # Load Investment data - NOTE: Sheet name has SPACES not underscores
        inv_df = pd.read_excel(excel_path, sheet_name='YTD Investment Activity')
        
        # Clean and standardize data
        ma_df = ma_df.fillna('Undisclosed')
        inv_df = inv_df.fillna('Undisclosed')
        
        return ma_df, inv_df
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        st.info("💡 Make sure your Excel file has sheets named 'YTD M&A Activity' and 'YTD Investment Activity' (with spaces)")
        return pd.DataFrame(), pd.DataFrame()

def save_data(ma_df, inv_df):
    """Save data back to Excel file with backup for undo"""
    try:
        # Try multiple possible file paths - INCLUDING data folder
        possible_paths = [
            'data/MedTech_YTD_Standardized.xlsx',
            './data/MedTech_YTD_Standardized.xlsx',
            'MedTech_YTD_Standardized.xlsx',
            'MedTech_MA_Masterlist.xlsx',
            './MedTech_MA_Masterlist.xlsx',
            '/mnt/project/MedTech_MA_Masterlist.xlsx',
            os.path.join(os.path.dirname(__file__), 'data', 'MedTech_YTD_Standardized.xlsx'),
            os.path.join(os.path.dirname(__file__), 'MedTech_YTD_Standardized.xlsx')
        ]
        
        excel_path = None
        for path in possible_paths:
            if os.path.exists(path):
                excel_path = path
                break
        
        if excel_path is None:
            # If file doesn't exist, create it in the data folder
            os.makedirs('data', exist_ok=True)
            excel_path = 'data/MedTech_YTD_Standardized.xlsx'
        
        # Create backup before saving (for undo functionality)
        backup_path = excel_path.replace('.xlsx', '_backup.xlsx')
        if os.path.exists(excel_path):
            import shutil
            shutil.copy2(excel_path, backup_path)
            st.session_state.last_backup_time = pd.Timestamp.now()
        
        # Save with correct sheet names (with spaces)
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
            ma_df.to_excel(writer, sheet_name='YTD M&A Activity', index=False)
            inv_df.to_excel(writer, sheet_name='YTD Investment Activity', index=False)
        
        st.session_state.changes_made = True
        return True
    except Exception as e:
        st.error(f"Error saving data: {str(e)}")
        st.warning("⚠️ Note: Streamlit Cloud has a read-only file system. Changes won't persist after app restarts.")
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
        
        # Restore from backup
        import shutil
        shutil.copy2(backup_path, excel_path)
        
        # Clear flags
        if 'changes_made' in st.session_state:
            del st.session_state.changes_made
        if 'last_backup_time' in st.session_state:
            del st.session_state.last_backup_time
        
        # Clear cache to reload data
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
    """Create quarterly stacked bar chart with deal count overlay - NO GRIDLINES"""
    try:
        # Prepare data
        quarterly_data = df.groupby('Quarter').agg({
            value_col: lambda x: sum([float(str(v).replace('$', '').replace('B', '').replace('M', '').replace(',', '')) 
                                     if v != 'Undisclosed' else 0 for v in x]),
            'Company': 'count'
        }).reset_index()
        quarterly_data.columns = ['Quarter', 'Total_Value', 'Deal_Count']
        
        # Sort quarters
        quarter_order = ['Q1', 'Q2', 'Q3', 'Q4']
        quarterly_data['Quarter'] = pd.Categorical(quarterly_data['Quarter'], categories=quarter_order, ordered=True)
        quarterly_data = quarterly_data.sort_values('Quarter')
        
        # Create figure
        fig = go.Figure()
        
        # Convert to millions for proper display
        quarterly_data['Total_Value_Millions'] = quarterly_data['Total_Value'] / 1000000
        
        # Add bar chart for deal values
        fig.add_trace(go.Bar(
            x=quarterly_data['Quarter'],
            y=quarterly_data['Total_Value_Millions'],
            name='Deal Value',
            marker_color='#7FA8C9',  # Muted blue
            text=[f"${v:,.0f}M" for v in quarterly_data['Total_Value_Millions']],  # Show in millions
            textposition='outside',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Deal Value: $%{y:,.0f}M<br><extra></extra>'
        ))
        
        # Add line chart for deal count
        fig.add_trace(go.Scatter(
            x=quarterly_data['Quarter'],
            y=quarterly_data['Deal_Count'],
            name='Deal Count',
            mode='lines+markers+text',
            line=dict(color='#C9A77F', width=3),  # Muted orange
            marker=dict(size=10),
            text=quarterly_data['Deal_Count'],
            textposition='top center',
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Deal Count: %{y}<br><extra></extra>'
        ))
        
        # Update layout - NO GRIDLINES
        fig.update_layout(
            title=title,
            xaxis=dict(title='Quarter', showgrid=False),
            yaxis=dict(
                title='Total Deal Value (USD Millions)',
                side='left',
                showgrid=False,  # Remove gridlines
                range=[0, max(quarterly_data['Total_Value_Millions']) * 1.2]  # Extend y-axis by 20% for data labels
            ),
            yaxis2=dict(
                title='Number of Deals',
                overlaying='y',
                side='right',
                showgrid=False,  # Remove gridlines
                range=[0, max(quarterly_data['Deal_Count']) * 1.3]  # Extend y2-axis by 30% for data labels
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
            margin=dict(t=100, b=50, l=50, r=50),
            plot_bgcolor='white',  # Clean white background
            paper_bgcolor='white'
        )
        
        return fig
    except Exception as e:
        st.error(f"Error creating chart: {str(e)}")
        return None

def load_jp_morgan_data(quarter):
    """Load JP Morgan data for a specific quarter from JSON file"""
    try:
        json_path = f'data/jp_morgan_2025_{quarter}.json'
        if os.path.exists(json_path):
            with open(json_path, 'r') as f:
                return json.load(f)
        return None
    except:
        return None

def create_jp_morgan_chart_by_category(category, color):
    """Create JP Morgan chart for a specific category with deal count overlay - NO GRIDLINES"""
    try:
        quarters = ['Q1', 'Q2', 'Q3']  # Only Q1-Q3, Q4 not available yet
        
        # Load data from JSON files or use defaults
        values = []
        counts = []
        
        for q in quarters:
            data = load_jp_morgan_data(q)
            if data:
                if category == 'M&A':
                    values.append(data['ma']['value'])
                    counts.append(data['ma']['count'])
                else:  # Venture
                    values.append(data['venture']['value'])
                    counts.append(data['venture']['count'])
            else:
                # Default data from JP Morgan 2025 reports if no JSON file exists
                data_map = {
                    'M&A': {
                        'values': [9200, 2100, 21700],  # Q1: $9.2B (57 deals), Q2: $2.1B (43 deals), Q3: $21.7B (65 deals)
                        'counts': [57, 43, 65]
                    },
                    'Venture': {
                        'values': [3700, 2600, 2900],  # Q1: $3.7B (117 rounds), Q2: $2.6B (90 rounds), Q3: $2.9B (67 rounds)
                        'counts': [117, 90, 67]
                    }
                }
                
                category_data = data_map.get(category, {'values': [0, 0, 0], 'counts': [0, 0, 0]})
                
                if q == 'Q1':
                    values.append(category_data['values'][0])
                    counts.append(category_data['counts'][0])
                elif q == 'Q2':
                    values.append(category_data['values'][1])
                    counts.append(category_data['counts'][1])
                else:  # Q3
                    values.append(category_data['values'][2])
                    counts.append(category_data['counts'][2])
        
        fig = go.Figure()
        
        # Add bars for deal values
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
        
        # Add line chart for deal count
        fig.add_trace(go.Scatter(
            x=quarters,
            y=counts,
            name='Deal Count',
            mode='lines+markers+text',
            line=dict(color='#8FB090', width=3),  # Muted green
            marker=dict(size=10, color='#8FB090'),
            text=[str(c) if c > 0 else '' for c in counts],
            textposition='top center',
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Deal Count: %{y}<br><extra></extra>'
        ))
        
        # Update layout with dual y-axes - NO GRIDLINES
        fig.update_layout(
            title=f'{category} Activity',
            xaxis=dict(title='Quarter', showgrid=False),
            yaxis=dict(
                title='Deal Value (USD Millions)',
                side='left',
                showgrid=False,  # Remove gridlines
                range=[0, max(values) * 1.2]  # Extend y-axis by 20% for data labels
            ),
            yaxis2=dict(
                title='Number of Deals',
                overlaying='y',
                side='right',
                showgrid=False,  # Remove gridlines
                range=[0, max(counts) * 1.3] if max(counts) > 0 else [0, 100]  # Extend y2-axis
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
            margin=dict(t=80, b=50, l=50, r=50),
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        
        return fig
    except Exception as e:
        st.error(f"Error creating {category} chart: {str(e)}")
        return None

# Main app
def main():
    st.title("🥼 MedTech M&A & Investment Activity Dashboard")
    
    # Load data
    ma_df, inv_df = load_data()
    
    # Sidebar navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Go to", ["Deal Activity", "JP Morgan Summary", "IPO Activity", "Data Management"])
    
    if page == "Deal Activity":
        show_deal_activity(ma_df, inv_df)
    elif page == "JP Morgan Summary":
        show_jp_morgan_summary(ma_df, inv_df)
    elif page == "IPO Activity":
        show_ipo_activity()
    elif page == "Data Management":
        show_data_management(ma_df, inv_df)

def show_deal_activity(ma_df, inv_df):
    """Display deal activity dashboard"""
    st.header("Deal Activity Dashboard")
    
    # M&A Activity Section - Full Width
    st.subheader("M&A Activity")
    
    # Search box
    search_ma = st.text_input("🔍 Search M&A Deals", placeholder="Search by company, acquirer, technology...", key='search_ma')
    
    # Filters
    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        quarters_ma = ['All'] + sorted(ma_df['Quarter'].unique().tolist())
        selected_quarter_ma = st.selectbox("Filter by Quarter", quarters_ma, key='ma_quarter')
    with filter_col2:
        months_ma = ['All'] + sorted(ma_df['Month'].unique().tolist())
        selected_month_ma = st.selectbox("Filter by Month", months_ma, key='ma_month')
    
    # Apply filters
    filtered_ma = ma_df.copy()
    if selected_quarter_ma != 'All':
        filtered_ma = filtered_ma[filtered_ma['Quarter'] == selected_quarter_ma]
    if selected_month_ma != 'All':
        filtered_ma = filtered_ma[filtered_ma['Month'] == selected_month_ma]
    
    # Apply search filter
    if search_ma:
        mask = filtered_ma.apply(lambda row: row.astype(str).str.contains(search_ma, case=False).any(), axis=1)
        filtered_ma = filtered_ma[mask]
    
    # Tabs for table, top deals, and charts
    tab1, tab2, tab3 = st.tabs(["📊 Table", "🏆 Top Deals", "📈 Charts"])
    
    with tab1:
        # Create display dataframe with sortable numeric values
        ma_display = filtered_ma.copy()
        
        # Add hidden numeric column for sorting - use -1 for Undisclosed so it goes to bottom
        def parse_to_numeric(val):
            if val == 'Undisclosed' or pd.isna(val):
                return -1  # Changed from 0 to -1 to sort Undisclosed to bottom
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                return float(val_str)
            except:
                return -1
        
        # Create a numeric sort column
        ma_display['_Deal_Value_Numeric'] = ma_display['Deal Value'].apply(parse_to_numeric)
        
        # Sort by Deal Value descending by default (highest deals first, Undisclosed at bottom)
        ma_display = ma_display.sort_values('_Deal_Value_Numeric', ascending=False)
        
        # Display without the numeric column (it's just for sorting)
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
        # Top 3 deals
        top_deals = filtered_ma.copy()
        
        # Parse function - values in Excel are already actual dollars like "$350,000,000"
        def parse_deal_value(val):
            if val == 'Undisclosed' or pd.isna(val):
                return 0
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                # Value is already in actual dollars, not millions
                return float(val_str)
            except:
                return 0
        
        top_deals['Deal_Value_Numeric'] = top_deals['Deal Value'].apply(parse_deal_value)
        top_deals = top_deals.nlargest(3, 'Deal_Value_Numeric')
        
        for idx, row in top_deals.iterrows():
            # Value is already in actual dollars, just format with commas
            formatted_value = str(row['Deal Value']) if row['Deal Value'] != 'Undisclosed' else 'Undisclosed'
            
            # Get deal type verb
            deal_type = row['Deal Type (Merger / Acquisition)']
            verb = "merged with" if deal_type == "Merger" else "acquired"
            
            # Display with value directly from Excel (already formatted) - muted color
            st.markdown(f"**{row['Acquirer']} {verb} {row['Company']}**")
            st.markdown(f"<h1 style='margin-top: -10px; margin-bottom: -10px; color: #7FA8C9;'>{formatted_value}</h1>", unsafe_allow_html=True)
            st.markdown("---")
    
    with tab3:
        fig = create_quarterly_chart(filtered_ma, 'Deal Value', 'M&A Activity by Quarter')
        if fig:
            st.plotly_chart(fig, use_container_width=True)
    
    # Add spacing between sections
    st.markdown("---")
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Venture Investment Activity Section - Full Width
    st.subheader("Venture Investment Activity")
    
    # Search box
    search_inv = st.text_input("🔍 Search Investment Deals", placeholder="Search by company, investors, technology...", key='search_inv')
    
    # Filters
    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        quarters_inv = ['All'] + sorted(inv_df['Quarter'].unique().tolist())
        selected_quarter_inv = st.selectbox("Filter by Quarter", quarters_inv, key='inv_quarter')
    with filter_col2:
        months_inv = ['All'] + sorted(inv_df['Month'].unique().tolist())
        selected_month_inv = st.selectbox("Filter by Month", months_inv, key='inv_month')
    
    # Apply filters
    filtered_inv = inv_df.copy()
    if selected_quarter_inv != 'All':
        filtered_inv = filtered_inv[filtered_inv['Quarter'] == selected_quarter_inv]
    if selected_month_inv != 'All':
        filtered_inv = filtered_inv[filtered_inv['Month'] == selected_month_inv]
    
    # Apply search filter
    if search_inv:
        mask = filtered_inv.apply(lambda row: row.astype(str).str.contains(search_inv, case=False).any(), axis=1)
        filtered_inv = filtered_inv[mask]
    
    # Tabs for table, top deals, and charts
    tab1, tab2, tab3 = st.tabs(["📊 Table", "🏆 Top Deals", "📈 Charts"])
    
    with tab1:
        # Format Amount Raised column for display with sortable numeric values
        inv_display = filtered_inv.copy()
        
        # Add numeric sort column - use -1 for Undisclosed so it goes to bottom
        inv_display['_Amount_Numeric'] = inv_display['Amount Raised'].apply(
            lambda x: float(x) if pd.notna(x) and x != 'Undisclosed' and str(x).replace('.','').replace('-','').isdigit() else -1
        )
        
        # Sort by Amount descending by default (highest amounts first, Undisclosed at bottom)
        inv_display = inv_display.sort_values('_Amount_Numeric', ascending=False)
        
        # Format for display
        inv_display['Amount Raised'] = inv_display['Amount Raised'].apply(
            lambda x: f"${x:,.0f}" if pd.notna(x) and x != 'Undisclosed' and str(x).replace('.','').replace('-','').isdigit() else x
        )
        
        # Display without the numeric column
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
        # Top 3 deals
        top_deals = filtered_inv.copy()
        
        # Parse function - values in Excel are already actual dollars like "$467,000,000"
        def parse_amount_value(val):
            if val == 'Undisclosed' or pd.isna(val):
                return 0
            val_str = str(val).replace('$', '').replace(',', '').strip()
            try:
                # Value is already in actual dollars, not millions
                return float(val_str)
            except:
                return 0
        
        top_deals['Amount_Numeric'] = top_deals['Amount Raised'].apply(parse_amount_value)
        top_deals = top_deals.nlargest(3, 'Amount_Numeric')
        
        for idx, row in top_deals.iterrows():
            # Format amount with commas
            amount_val = row['Amount Raised']
            if pd.notna(amount_val) and amount_val != 'Undisclosed':
                try:
                    formatted_value = f"${float(amount_val):,.0f}"
                except:
                    formatted_value = str(amount_val)
            else:
                formatted_value = "Undisclosed"
            
            # Display with formatted value - muted color
            st.markdown(f"**{row['Company']}**")
            st.markdown(f"<h1 style='margin-top: -10px; margin-bottom: -10px; color: #C9A77F;'>{formatted_value}</h1>", unsafe_allow_html=True)
            st.markdown("---")
    
    with tab3:
        fig = create_quarterly_chart(filtered_inv, 'Amount Raised', 'Venture Investment by Quarter')
        if fig:
            st.plotly_chart(fig, use_container_width=True)

def show_jp_morgan_summary(ma_df, inv_df):
    """Display JP Morgan summary"""
    st.header("JP Morgan MedTech Industry Report")
    
    # Calculate BeaconOne quarterly stats
    def calc_quarterly_stats(df, quarter, value_col):
        q_data = df[df['Quarter'] == quarter]
        
        # Parse values
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
        
        # Format value
        if total_value >= 1000000000:
            formatted_value = f"${total_value/1000000000:.1f}B"
        elif total_value >= 1000000:
            formatted_value = f"${total_value/1000000:.0f}M"
        else:
            formatted_value = "$0"
            
        return count, formatted_value, total_value
    
    # Calculate stats for each quarter
    beacon_stats = {}
    for q in ['Q1', 'Q2', 'Q3']:
        ma_count, ma_value, ma_value_raw = calc_quarterly_stats(ma_df, q, 'Deal Value')
        inv_count, inv_value, inv_value_raw = calc_quarterly_stats(inv_df, q, 'Amount Raised')
        beacon_stats[q] = {
            'ma_count': ma_count,
            'ma_value': ma_value,
            'ma_value_raw': ma_value_raw,
            'inv_count': inv_count,
            'inv_value': inv_value,
            'inv_value_raw': inv_value_raw
        }
    
    st.markdown("### 2025 Q1-Q3 Activity by Category")
    
    # Create 1x2 grid for charts (only M&A and Venture)
    col1, col2 = st.columns(2)
    
    # Left: M&A
    with col1:
        fig_ma = create_jp_morgan_chart_by_category('M&A', '#7FA8C9')  # Muted blue
        if fig_ma:
            st.plotly_chart(fig_ma, use_container_width=True)
    
    # Right: Venture
    with col2:
        fig_venture = create_jp_morgan_chart_by_category('Venture', '#C9A77F')  # Muted orange
        if fig_venture:
            st.plotly_chart(fig_venture, use_container_width=True)
    
    # Key trends below the charts
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
        st.markdown("**Overarching Trend**: Medtech M&A activity increased through Q3 2025, surpassing full-year 2024 numbers, with strategic consolidation driving large-scale transactions")
        
    with col2:
        st.markdown("**Venture Capital**")
        st.markdown("")
        st.markdown("• **Q1 2025**: Medtech venture investment activity continued to see larger rounds into fewer companies to post a higher dollar total for Q1 2025, exceeding Q1 2024")
        st.markdown("")
        st.markdown("• **Q2 2025**: The medtech venture landscape continues to show resilience, with total venture funding reaching $6.8 billion in the first half of 2025, positioning the sector to potentially exceed 2024's $12.7 billion full-year total")
        st.markdown("")
        st.markdown("• **Q3 2025**: Medtech venture funding started the year strong yet had a weaker Q2 and Q3 in a challenging venture funding environment across all of healthcare and life sciences")
        st.markdown("")
        st.markdown("**Overarching Trend**: Late-stage venture rounds continue to dominate at $7.9B YTD, while early-stage funding remains selective as investors focus on companies with proven traction")

    # Add YTD comparison section
    st.markdown("---")
    st.markdown("### JPMorgan vs BeaconOne Data - YTD 2025 Comparison")
    
    # Load JP Morgan data for comparison
    jp_data = {}
    for q in ['Q1', 'Q2', 'Q3']:
        loaded_data = load_jp_morgan_data(q)
        if loaded_data:
            jp_data[q] = loaded_data
        else:
            # Use default values
            jp_data[q] = {
                'Q1': {
                    'ma': {'value': 9200, 'count': 57},
                    'venture': {'value': 3700, 'count': 117}
                },
                'Q2': {
                    'ma': {'value': 2100, 'count': 43},
                    'venture': {'value': 2600, 'count': 90}
                },
                'Q3': {
                    'ma': {'value': 21700, 'count': 65},
                    'venture': {'value': 2900, 'count': 67}
                }
            }[q]
    
    # Calculate YTD totals
    jp_ma_ytd_value = sum([jp_data['Q1']['ma']['value'], jp_data['Q2']['ma']['value'], jp_data['Q3']['ma']['value']])
    jp_ma_ytd_count = sum([jp_data['Q1']['ma']['count'], jp_data['Q2']['ma']['count'], jp_data['Q3']['ma']['count']])
    jp_vc_ytd_value = sum([jp_data['Q1']['venture']['value'], jp_data['Q2']['venture']['value'], jp_data['Q3']['venture']['value']])
    jp_vc_ytd_count = sum([jp_data['Q1']['venture']['count'], jp_data['Q2']['venture']['count'], jp_data['Q3']['venture']['count']])
    
    beacon_ma_ytd_value = sum([beacon_stats['Q1']['ma_value_raw'], beacon_stats['Q2']['ma_value_raw'], beacon_stats['Q3']['ma_value_raw']]) / 1000000
    beacon_ma_ytd_count = sum([beacon_stats['Q1']['ma_count'], beacon_stats['Q2']['ma_count'], beacon_stats['Q3']['ma_count']])
    beacon_vc_ytd_value = sum([beacon_stats['Q1']['inv_value_raw'], beacon_stats['Q2']['inv_value_raw'], beacon_stats['Q3']['inv_value_raw']]) / 1000000
    beacon_vc_ytd_count = sum([beacon_stats['Q1']['inv_count'], beacon_stats['Q2']['inv_count'], beacon_stats['Q3']['inv_count']])
    
    # Create two columns for M&A and Venture comparison charts
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### M&A Activity - YTD Comparison")
        
        # Create simplified comparison chart for M&A (deal value only)
        fig_ma_comp = go.Figure()
        
        # JPMorgan bar
        fig_ma_comp.add_trace(go.Bar(
            name='JPMorgan',
            x=['Deal Value'],
            y=[jp_ma_ytd_value / 1000],  # Convert to billions for display
            marker_color='#7FA8C9',
            text=[f'${jp_ma_ytd_value/1000:.1f}B'],
            textposition='outside',
            width=0.4,
            hovertemplate='<b>JPMorgan</b><br>Deal Value: $%{y:.1f}B<extra></extra>'
        ))
        
        # BeaconOne bar
        fig_ma_comp.add_trace(go.Bar(
            name='BeaconOne',
            x=['Deal Value'],
            y=[beacon_ma_ytd_value / 1000],  # Convert to billions for display
            marker_color='#A8C9D1',
            text=[f'${beacon_ma_ytd_value/1000:.1f}B'],
            textposition='outside',
            width=0.4,
            hovertemplate='<b>BeaconOne</b><br>Deal Value: $%{y:.1f}B<extra></extra>'
        ))
        
        fig_ma_comp.update_layout(
            barmode='group',
            height=400,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            xaxis=dict(showgrid=False, title=''),
            yaxis=dict(
                title='Deal Value (USD Billions)',
                showgrid=False,
                range=[0, max(jp_ma_ytd_value, beacon_ma_ytd_value) / 1000 * 1.2]
            ),
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(t=60, b=50, l=50, r=50)
        )
        
        st.plotly_chart(fig_ma_comp, use_container_width=True)
        
        # Summary metrics below chart with both value and count
        st.markdown(f"""
        <div style='display: flex; justify-content: space-around; margin-top: 15px;'>
            <div style='text-align: center; background-color: #F0F4F7; padding: 20px; border-radius: 8px; flex: 1; margin: 0 5px;'>
                <p style='margin: 0; font-size: 12px; color: #666; font-weight: 500;'>JPMorgan YTD</p>
                <p style='margin: 8px 0 0 0; font-size: 32px; font-weight: bold; color: #7FA8C9;'>${jp_ma_ytd_value/1000:.1f}B</p>
                <p style='margin: 5px 0 0 0; font-size: 16px; color: #888;'>{jp_ma_ytd_count} deals</p>
            </div>
            <div style='text-align: center; background-color: #F0F4F7; padding: 20px; border-radius: 8px; flex: 1; margin: 0 5px;'>
                <p style='margin: 0; font-size: 12px; color: #666; font-weight: 500;'>BeaconOne YTD</p>
                <p style='margin: 8px 0 0 0; font-size: 32px; font-weight: bold; color: #A8C9D1;'>${beacon_ma_ytd_value/1000:.1f}B</p>
                <p style='margin: 5px 0 0 0; font-size: 16px; color: #888;'>{int(beacon_ma_ytd_count)} deals</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("#### Venture Investment - YTD Comparison")
        
        # Create simplified comparison chart for Venture (deal value only)
        fig_vc_comp = go.Figure()
        
        # JPMorgan bar
        fig_vc_comp.add_trace(go.Bar(
            name='JPMorgan',
            x=['Deal Value'],
            y=[jp_vc_ytd_value / 1000],  # Convert to billions for display
            marker_color='#C9A77F',
            text=[f'${jp_vc_ytd_value/1000:.1f}B'],
            textposition='outside',
            width=0.4,
            hovertemplate='<b>JPMorgan</b><br>Deal Value: $%{y:.1f}B<extra></extra>'
        ))
        
        # BeaconOne bar
        fig_vc_comp.add_trace(go.Bar(
            name='BeaconOne',
            x=['Deal Value'],
            y=[beacon_vc_ytd_value / 1000],  # Convert to billions for display
            marker_color='#D9C9A8',
            text=[f'${beacon_vc_ytd_value/1000:.1f}B'],
            textposition='outside',
            width=0.4,
            hovertemplate='<b>BeaconOne</b><br>Deal Value: $%{y:.1f}B<extra></extra>'
        ))
        
        fig_vc_comp.update_layout(
            barmode='group',
            height=400,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            xaxis=dict(showgrid=False, title=''),
            yaxis=dict(
                title='Deal Value (USD Billions)',
                showgrid=False,
                range=[0, max(jp_vc_ytd_value, beacon_vc_ytd_value) / 1000 * 1.2]
            ),
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(t=60, b=50, l=50, r=50)
        )
        
        st.plotly_chart(fig_vc_comp, use_container_width=True)
        
        # Summary metrics below chart with both value and count
        st.markdown(f"""
        <div style='display: flex; justify-content: space-around; margin-top: 15px;'>
            <div style='text-align: center; background-color: #F9F7F4; padding: 20px; border-radius: 8px; flex: 1; margin: 0 5px;'>
                <p style='margin: 0; font-size: 12px; color: #666; font-weight: 500;'>JPMorgan YTD</p>
                <p style='margin: 8px 0 0 0; font-size: 32px; font-weight: bold; color: #C9A77F;'>${jp_vc_ytd_value/1000:.1f}B</p>
                <p style='margin: 5px 0 0 0; font-size: 16px; color: #888;'>{jp_vc_ytd_count} deals</p>
            </div>
            <div style='text-align: center; background-color: #F9F7F4; padding: 20px; border-radius: 8px; flex: 1; margin: 0 5px;'>
                <p style='margin: 0; font-size: 12px; color: #666; font-weight: 500;'>BeaconOne YTD</p>
                <p style='margin: 8px 0 0 0; font-size: 32px; font-weight: bold; color: #D9C9A8;'>${beacon_vc_ytd_value/1000:.1f}B</p>
                <p style='margin: 5px 0 0 0; font-size: 16px; color: #888;'>{int(beacon_vc_ytd_count)} deals</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

def show_ipo_activity():
    """Display IPO activity from JP Morgan reports"""
    st.header("IPO Activity")
    
    # Load IPO data from JSON files
    ipo_data = {}
    for q in ['Q1', 'Q2', 'Q3', 'Q4']:
        json_path = f'data/jp_morgan_ipo_{q}_2025.json'
        if os.path.exists(json_path):
            with open(json_path, 'r') as f:
                ipo_data[q] = json.load(f)
    
    # If no data exists, show default/example data
    if not ipo_data:
        st.info("📊 No IPO data has been uploaded yet. Use the Data Management page to add IPO information from JP Morgan reports.")
        
        # Show example structure
        st.markdown("### Example IPO Data Structure")
        st.markdown("""
        When you upload IPO data, you'll see:
        - **YTD Chart**: Visual representation of IPO activity by quarter
        - **IPO List**: Detailed information about each IPO including:
          - Company name
          - IPO value
          - Date
          - Key details from the report
        """)
        return
    
    # Calculate YTD totals
    quarters = []
    values = []
    counts = []
    
    for q in ['Q1', 'Q2', 'Q3', 'Q4']:
        if q in ipo_data:
            quarters.append(q)
            values.append(ipo_data[q].get('total_value', 0))
            counts.append(ipo_data[q].get('count', 0))
    
    # Create YTD chart
    if quarters:
        st.markdown("### YTD 2025 IPO Activity")
        
        fig = go.Figure()
        
        # Add bar chart for IPO values
        fig.add_trace(go.Bar(
            x=quarters,
            y=values,
            name='IPO Value',
            marker_color='#9B8FAC',  # Muted purple
            text=[f'${v:,.0f}M' if v > 0 else 'No IPOs' for v in values],
            textposition='outside',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>IPO Value: $%{y:,.0f}M<br><extra></extra>'
        ))
        
        # Add line chart for IPO count
        fig.add_trace(go.Scatter(
            x=quarters,
            y=counts,
            name='IPO Count',
            mode='lines+markers+text',
            line=dict(color='#8A7A98', width=3),
            marker=dict(size=10),
            text=counts,
            textposition='top center',
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>IPO Count: %{y}<br><extra></extra>'
        ))
        
        fig.update_layout(
            xaxis=dict(title='Quarter', showgrid=False),
            yaxis=dict(
                title='Total IPO Value (USD Millions)',
                side='left',
                showgrid=False,
                range=[0, max(values) * 1.3] if max(values) > 0 else [0, 100]
            ),
            yaxis2=dict(
                title='Number of IPOs',
                overlaying='y',
                side='right',
                showgrid=False,
                range=[0, max(counts) * 1.4] if max(counts) > 0 else [0, 10]
            ),
            hovermode='x unified',
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=450,
            margin=dict(t=60, b=50, l=50, r=50),
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # YTD Summary
        total_value = sum(values)
        total_count = sum(counts)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total YTD IPO Value", f"${total_value:,.0f}M")
        with col2:
            st.metric("Total YTD IPOs", total_count)
        with col3:
            avg_value = total_value / total_count if total_count > 0 else 0
            st.metric("Average IPO Size", f"${avg_value:,.0f}M")
    
    # Display detailed IPO list
    st.markdown("---")
    st.markdown("### 2025 IPO Details")
    
    # Collect all IPOs
    all_ipos = []
    for q in ['Q1', 'Q2', 'Q3', 'Q4']:
        if q in ipo_data and 'ipos' in ipo_data[q]:
            for ipo in ipo_data[q]['ipos']:
                ipo['quarter'] = q
                all_ipos.append(ipo)
    
    if all_ipos:
        # Sort by value (descending)
        all_ipos.sort(key=lambda x: x.get('value', 0), reverse=True)
        
        # Display each IPO
        for ipo in all_ipos:
            company = ipo.get('company', 'Unknown Company')
            value = ipo.get('value', 0)
            quarter = ipo.get('quarter', 'Q1')
            date = ipo.get('date', 'Date not provided')
            details = ipo.get('details', [])
            
            # Create expandable section for each IPO
            with st.expander(f"**{company}** - ${value:,.0f}M ({quarter} 2025)", expanded=False):
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    st.markdown(f"**IPO Value:** ${value:,.0f}M")
                    st.markdown(f"**Quarter:** {quarter} 2025")
                    st.markdown(f"**Date:** {date}")
                
                with col2:
                    if details:
                        st.markdown("**Key Details:**")
                        for detail in details:
                            st.markdown(f"• {detail}")
                    else:
                        st.markdown("*No additional details provided*")
    else:
        st.info("No IPO details available. Add IPO information through the Data Management page.")

def show_data_management(ma_df, inv_df):
    """Data management page for adding deals and uploading JP Morgan reports"""
    st.header("Data Management")
    
    # Add undo button at the top
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
    
    # Create tabs for different data management tasks
    tab1, tab2, tab3 = st.tabs(["📝 Add Manual Deals", "📊 Upload JP Morgan Report", "🎯 Add IPO Data"])
    
    with tab1:
        show_manual_deal_entry(ma_df, inv_df)
    
    with tab2:
        show_jp_morgan_upload()
    
    with tab3:
        show_ipo_data_upload()

def show_manual_deal_entry(ma_df, inv_df):
    """Manual deal entry forms"""
    st.subheader("Add New Deal Manually")
    
    # Select deal type
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
                    # Parse deal value to standardized format
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
                    
                    # Append to dataframe
                    ma_df_updated = pd.concat([ma_df, new_deal], ignore_index=True)
                    
                    # Save data
                    if save_data(ma_df_updated, inv_df):
                        st.success("✅ M&A deal added successfully!")
                        st.balloons()
                        # Clear cache to reload data
                        st.cache_data.clear()
                        st.rerun()
                else:
                    st.error("Please fill in all required fields (*)")
    
    else:  # Venture Investment
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
                    # Parse amount to numeric format (just the number, no formatting)
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
                    
                    # Append to dataframe
                    inv_df_updated = pd.concat([inv_df, new_deal], ignore_index=True)
                    
                    # Save data
                    if save_data(ma_df, inv_df_updated):
                        st.success("✅ Investment deal added successfully!")
                        st.balloons()
                        # Clear cache to reload data
                        st.cache_data.clear()
                        st.rerun()
                else:
                    st.error("Please fill in all required fields (*)")

def show_jp_morgan_upload():
    """JP Morgan report upload and data extraction"""
    st.subheader("Upload JP Morgan MedTech Industry Report")
    
    st.info("""
    📄 **Instructions:**
    1. Upload the quarterly JP Morgan MedTech Industry Report (PDF or text)
    2. Enter the key metrics manually below
    3. Charts and comparison cards will automatically update
    """)
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Choose JP Morgan Report (Optional)", 
        type=['pdf', 'txt', 'docx'],
        help="Upload the quarterly JP Morgan MedTech Industry Report"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    # Quarter selection
    col1, col2 = st.columns(2)
    with col1:
        report_year = st.selectbox("Report Year", [2025, 2024, 2023])
    with col2:
        report_quarter = st.selectbox("Report Quarter", ["Q1", "Q2", "Q3", "Q4"])
    
    st.markdown("### Enter Data Manually")
    st.markdown("Please enter the key metrics from the report:")
    
    with st.form("jp_morgan_data_form"):
        st.markdown("#### M&A Activity")
        col1, col2 = st.columns(2)
        with col1:
            ma_value = st.number_input("M&A Deal Value ($M)", min_value=0.0, value=0.0, step=100.0)
        with col2:
            ma_count = st.number_input("M&A Deal Count", min_value=0, value=0, step=1)
        
        st.markdown("#### Venture Capital")
        col1, col2 = st.columns(2)
        with col1:
            vc_value = st.number_input("Venture Deal Value ($M)", min_value=0.0, value=0.0, step=100.0)
        with col2:
            vc_count = st.number_input("Venture Deal Count", min_value=0, value=0, step=1)
        
        st.markdown("#### Key Takeaways (Optional)")
        ma_takeaway = st.text_area("M&A Key Takeaway", placeholder="Enter key insight for M&A activity...")
        vc_takeaway = st.text_area("Venture Capital Key Takeaway", placeholder="Enter key insight for VC activity...")
        
        submitted = st.form_submit_button("💾 Save JP Morgan Data")
        
        if submitted:
            # Save the data to a JSON file
            jp_morgan_data = {
                'year': report_year,
                'quarter': report_quarter,
                'ma': {'value': ma_value, 'count': ma_count, 'takeaway': ma_takeaway},
                'venture': {'value': vc_value, 'count': vc_count, 'takeaway': vc_takeaway}
            }
            
            # Create data directory if it doesn't exist
            os.makedirs('data', exist_ok=True)
            
            # Save to JSON file
            json_path = f'data/jp_morgan_{report_year}_{report_quarter}.json'
            with open(json_path, 'w') as f:
                json.dump(jp_morgan_data, f, indent=2)
            
            st.success(f"✅ JP Morgan {report_year} {report_quarter} data saved successfully!")
            st.info("📊 The JP Morgan Summary page will now reflect this data. Navigate to 'JP Morgan Summary' to view the updated charts and comparison cards.")
            st.balloons()
            
            # Clear cache to reload data
            st.cache_data.clear()

def show_ipo_data_upload():
    """Upload and manage IPO data from JP Morgan reports"""
    st.subheader("Add IPO Data from JP Morgan Reports")
    
    st.info("""
    📈 **Instructions:**
    1. Select the quarter for the IPO data
    2. Enter aggregate IPO statistics (total value and count)
    3. Add individual IPO details with company information
    """)
    
    # Quarter selection
    col1, col2 = st.columns(2)
    with col1:
        report_year = st.selectbox("Report Year", [2025, 2024, 2023], key='ipo_year')
    with col2:
        report_quarter = st.selectbox("Report Quarter", ["Q1", "Q2", "Q3", "Q4"], key='ipo_quarter')
    
    st.markdown("### Quarterly IPO Summary")
    
    with st.form("ipo_summary_form"):
        col1, col2 = st.columns(2)
        with col1:
            total_ipo_value = st.number_input("Total IPO Value ($M)", min_value=0.0, value=0.0, step=10.0)
        with col2:
            total_ipo_count = st.number_input("Total Number of IPOs", min_value=0, value=0, step=1)
        
        st.markdown("### Individual IPO Details")
        st.markdown("Add information about specific IPOs in this quarter:")
        
        # Number of IPOs to add
        num_ipos = st.number_input("How many IPOs to add?", min_value=0, max_value=20, value=0, step=1)
        
        ipos_list = []
        
        for i in range(int(num_ipos)):
            st.markdown(f"#### IPO #{i+1}")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                company_name = st.text_input(f"Company Name", key=f"ipo_company_{i}")
            with col2:
                ipo_value = st.number_input(f"IPO Value ($M)", min_value=0.0, value=0.0, step=1.0, key=f"ipo_value_{i}")
            with col3:
                ipo_date = st.text_input(f"Date", placeholder="e.g., Jan 15, 2025", key=f"ipo_date_{i}")
            
            # Details as bullet points
            ipo_details = st.text_area(
                f"Key Details (one per line)",
                placeholder="Enter each detail on a new line:\n- Detail 1\n- Detail 2\n- Detail 3",
                key=f"ipo_details_{i}",
                height=100
            )
            
            if company_name:
                # Parse details into list
                details_list = []
                if ipo_details:
                    for line in ipo_details.split('\n'):
                        line = line.strip()
                        if line:
                            # Remove leading bullet points or dashes
                            line = line.lstrip('•-– ').strip()
                            if line:
                                details_list.append(line)
                
                ipos_list.append({
                    'company': company_name,
                    'value': ipo_value,
                    'date': ipo_date if ipo_date else 'Date not provided',
                    'details': details_list
                })
        
        submitted = st.form_submit_button("💾 Save IPO Data")
        
        if submitted:
            # Create the IPO data structure
            ipo_data = {
                'year': report_year,
                'quarter': report_quarter,
                'total_value': total_ipo_value,
                'count': total_ipo_count,
                'ipos': ipos_list
            }
            
            # Create data directory if it doesn't exist
            os.makedirs('data', exist_ok=True)
            
            # Save to JSON file
            json_path = f'data/jp_morgan_ipo_{report_quarter}_{report_year}.json'
            with open(json_path, 'w') as f:
                json.dump(ipo_data, f, indent=2)
            
            st.success(f"✅ IPO data for {report_quarter} {report_year} saved successfully!")
            st.info("📊 Navigate to 'IPO Activity' page to view the updated charts and details.")
            st.balloons()
            
            # Clear cache
            st.cache_data.clear()

if __name__ == "__main__":
    main()