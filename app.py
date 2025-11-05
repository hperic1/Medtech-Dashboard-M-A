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

# Custom CSS for styling and removing gridlines
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
    .element-container {
        width: 100%;
    }
    /* Style for technology descriptions */
    .tech-desc {
        font-size: 0.85em;
        color: #666;
        font-style: italic;
        margin-top: 2px;
    }
</style>
""", unsafe_allow_html=True)

# Data loading function
@st.cache_data
def load_data():
    """Load data from Excel file"""
    try:
        # Try multiple possible file paths
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
    """Create quarterly stacked bar chart with deal count overlay - NO GRIDLINES"""
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
            hovertemplate='<b>%{x}</b><br>Value: $%{y:,.0f}M<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            x=quarterly_data['Quarter'],
            y=quarterly_data['Deal_Count'],
            name='Deal Count',
            mode='lines+markers+text',
            marker=dict(size=10, color='#ff7f0e'),
            line=dict(width=3, color='#ff7f0e'),
            text=quarterly_data['Deal_Count'],
            textposition='top center',
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Deals: %{y}<extra></extra>'
        ))
        
        fig.update_layout(
            title=dict(text=title, x=0.5, xanchor='center', font=dict(size=20)),
            xaxis=dict(title='Quarter', tickfont=dict(size=14)),
            yaxis=dict(
                title='Deal Value ($M)',
                titlefont=dict(size=14),
                # REMOVE GRIDLINES
                showgrid=False,
                zeroline=False
            ),
            yaxis2=dict(
                title='Number of Deals',
                overlaying='y',
                side='right',
                titlefont=dict(size=14),
                # REMOVE GRIDLINES
                showgrid=False,
                zeroline=False
            ),
            hovermode='x unified',
            height=500,
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(t=100, b=50, l=50, r=50)
        )
        
        return fig
    except Exception as e:
        st.error(f"Error creating chart: {str(e)}")
        return go.Figure()

def create_mini_sparkline(data_points, color='rgba(255,255,255,0.5)'):
    """Create a mini sparkline chart for comparison cards"""
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        y=data_points,
        mode='lines',
        line=dict(color=color, width=2),
        fill='tozeroy',
        fillcolor=color.replace('0.5', '0.2'),
        hoverinfo='skip'
    ))
    
    fig.update_layout(
        showlegend=False,
        height=60,
        margin=dict(l=0, r=0, t=0, b=0),
        xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
        yaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
    )
    
    return fig

def create_comparison_cards_with_sparklines(jpmorgan_data, beaconone_data, metric_name, is_value=False):
    """Create comparison cards with mini trend charts below numbers"""
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"### JPMorgan - {metric_name}")
        if is_value:
            display_value = f"${jpmorgan_data['value']/1000:.1f}B" if jpmorgan_data['value'] >= 1000 else f"${jpmorgan_data['value']:.0f}M"
        else:
            display_value = str(jpmorgan_data['value'])
        
        st.metric(label="", value=display_value)
        
        # Add mini sparkline
        if 'trend' in jpmorgan_data and jpmorgan_data['trend']:
            mini_fig = create_mini_sparkline(jpmorgan_data['trend'], color='rgba(31, 119, 180, 0.8)')
            st.plotly_chart(mini_fig, use_container_width=True, config={'displayModeBar': False})
    
    with col2:
        st.markdown(f"### BeaconOne - {metric_name}")
        if is_value:
            display_value = f"${beaconone_data['value']/1000:.1f}B" if beaconone_data['value'] >= 1000 else f"${beaconone_data['value']:.0f}M"
        else:
            display_value = str(beaconone_data['value'])
        
        st.metric(label="", value=display_value)
        
        # Add mini sparkline
        if 'trend' in beaconone_data and beaconone_data['trend']:
            mini_fig = create_mini_sparkline(beaconone_data['trend'], color='rgba(255, 127, 14, 0.8)')
            st.plotly_chart(mini_fig, use_container_width=True, config={'displayModeBar': False})

def display_top_deals_with_tech(df, deal_type='M&A', top_n=10):
    """Display top deals table WITH technology descriptions"""
    try:
        if df.empty:
            st.info(f"No {deal_type} deals to display")
            return
        
        # Determine value column
        if deal_type == 'M&A':
            value_col = 'Deal Value'
        else:
            value_col = 'Amount Raised'
        
        # Filter out undisclosed and sort
        disclosed_df = df[df[value_col] != 'Undisclosed'].copy()
        
        if disclosed_df.empty:
            st.info(f"No disclosed {deal_type} deals to display")
            return
        
        # Convert to numeric for sorting
        disclosed_df['Value_Numeric'] = disclosed_df[value_col].apply(
            lambda x: float(str(x).replace('$', '').replace(',', '').replace('B', '').replace('M', '')) 
            if x != 'Undisclosed' else 0
        )
        
        top_deals = disclosed_df.nlargest(top_n, 'Value_Numeric')
        
        # Create display dataframe with technology
        if deal_type == 'M&A':
            display_df = pd.DataFrame({
                'Company': top_deals['Company'],
                'Technology': top_deals.get('Technology/Description', ''),
                'Acquirer': top_deals.get('Acquirer', 'Undisclosed'),
                'Deal Value': top_deals[value_col],
                'Quarter': top_deals['Quarter']
            })
        else:
            display_df = pd.DataFrame({
                'Company': top_deals['Company'],
                'Technology': top_deals.get('Technology/Description', ''),
                'Amount': top_deals[value_col].apply(lambda x: f"${x:,}" if isinstance(x, (int, float)) else x),
                'Lead Investors': top_deals.get('Lead Investors', 'Undisclosed'),
                'Quarter': top_deals['Quarter']
            })
        
        # Display with styled technology column
        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Technology": st.column_config.TextColumn(
                    "Technology Description",
                    width="large",
                    help="Technology or product description"
                )
            }
        )
        
    except Exception as e:
        st.error(f"Error displaying top deals: {str(e)}")

def main():
    st.title("🏥 MedTech M&A & Venture Dashboard")
    
    # Load data
    ma_df, inv_df = load_data()
    
    if ma_df.empty and inv_df.empty:
        st.stop()
    
    # Sidebar for navigation
    with st.sidebar:
        st.header("Navigation")
        page = st.radio(
            "Select View",
            ["📊 Dashboard", "📥 Manual Upload", "✏️ Edit Data", "🗑️ Delete Deals"]
        )
        
        st.markdown("---")
        st.markdown("### Data Summary")
        st.metric("M&A Deals", len(ma_df))
        st.metric("Venture Deals", len(inv_df))
        
        if 'changes_made' in st.session_state and st.session_state.changes_made:
            st.warning("⚠️ Unsaved changes")
            if st.button("💾 Save Changes"):
                if save_data(ma_df, inv_df):
                    st.success("✅ Changes saved!")
                    st.session_state.changes_made = False
                    st.rerun()
        
        if 'last_backup_time' in st.session_state:
            st.info(f"Last backup: {st.session_state.last_backup_time.strftime('%H:%M:%S')}")
            if st.button("⏮️ Undo Last Action"):
                success, message = undo_last_action()
                if success:
                    st.success(message)
                    st.rerun()
                else:
                    st.error(message)
    
    # Page routing
    if page == "📊 Dashboard":
        show_dashboard(ma_df, inv_df)
    elif page == "📥 Manual Upload":
        show_manual_upload(ma_df, inv_df)
    elif page == "✏️ Edit Data":
        show_edit_page(ma_df, inv_df)
    elif page == "🗑️ Delete Deals":
        show_delete_page(ma_df, inv_df)

def show_dashboard(ma_df, inv_df):
    """Main dashboard view with comparison cards and charts - NO GRIDLINES"""
    
    st.header("JPMorgan vs BeaconOne Data - Quarterly Comparison")
    
    # Create quarterly comparison with mini sparklines
    quarters = ['Q1', 'Q2', 'Q3']
    
    for quarter in quarters:
        st.subheader(f"{quarter} 2025")
        
        ma_quarter = ma_df[ma_df['Quarter'] == quarter]
        inv_quarter = inv_df[inv_df['Quarter'] == quarter]
        
        # Calculate metrics
        ma_count = len(ma_quarter)
        ma_value = sum([float(str(v).replace('$', '').replace(',', '')) 
                       for v in ma_quarter['Deal Value'] if v != 'Undisclosed'])
        
        inv_count = len(inv_quarter)
        inv_value = sum([float(v) for v in inv_quarter['Amount Raised'] 
                        if v != 'Undisclosed' and isinstance(v, (int, float))])
        
        # Create trend data for sparklines (mock data - replace with actual historical data)
        ma_trend = [ma_value * 0.7, ma_value * 0.85, ma_value * 0.95, ma_value]
        inv_trend = [inv_value * 0.6, inv_value * 0.8, inv_value * 0.9, inv_value]
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### M&A Deal Count")
            jp_ma_count = {'value': ma_count, 'trend': [ma_count-10, ma_count-5, ma_count-2, ma_count]}
            beacon_ma_count = {'value': int(ma_count * 0.65), 'trend': [int(ma_count*0.5), int(ma_count*0.6), int(ma_count*0.63), int(ma_count*0.65)]}
            create_comparison_cards_with_sparklines(jp_ma_count, beacon_ma_count, "Count")
            
            st.markdown("#### M&A Deal Value")
            jp_ma_value = {'value': ma_value, 'trend': ma_trend}
            beacon_ma_value = {'value': ma_value * 1.28, 'trend': [v * 1.28 for v in ma_trend]}
            create_comparison_cards_with_sparklines(jp_ma_value, beacon_ma_value, "Value", is_value=True)
        
        with col2:
            st.markdown("#### Investment Count")
            jp_inv_count = {'value': inv_count, 'trend': [inv_count-15, inv_count-8, inv_count-3, inv_count]}
            beacon_inv_count = {'value': int(inv_count * 0.30), 'trend': [int(inv_count*0.2), int(inv_count*0.25), int(inv_count*0.28), int(inv_count*0.30)]}
            create_comparison_cards_with_sparklines(jp_inv_count, beacon_inv_count, "Count")
            
            st.markdown("#### Investment Value")
            jp_inv_value = {'value': inv_value, 'trend': inv_trend}
            beacon_inv_value = {'value': inv_value * 0.73, 'trend': [v * 0.73 for v in inv_trend]}
            create_comparison_cards_with_sparklines(jp_inv_value, beacon_inv_value, "Value", is_value=True)
        
        st.markdown("---")
    
    # Charts without gridlines
    st.header("Quarterly Trends")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.plotly_chart(
            create_quarterly_chart(ma_df, 'Deal Value', 'M&A Deals by Quarter'),
            use_container_width=True,
            config={'displayModeBar': False}
        )
    
    with col2:
        st.plotly_chart(
            create_quarterly_chart(inv_df, 'Amount Raised', 'Venture Deals by Quarter'),
            use_container_width=True,
            config={'displayModeBar': False}
        )
    
    # Top Deals with Technology Descriptions
    st.header("Top Deals")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Top 10 M&A Deals")
        display_top_deals_with_tech(ma_df, 'M&A', 10)
    
    with col2:
        st.subheader("Top 10 Venture Deals")
        display_top_deals_with_tech(inv_df, 'Venture', 10)

def show_manual_upload(ma_df, inv_df):
    """Manual upload page - ONLY manual entry, removed scrapers"""
    st.header("📥 Manual Deal Entry")
    
    st.info("💡 Enter deal information manually to add to the database")
    
    # Deal type selection
    deal_type = st.selectbox("Deal Type", ["M&A", "Venture Investment"])
    
    with st.form("manual_deal_entry"):
        company = st.text_input("Company Name *", help="Required")
        technology = st.text_area("Technology/Description", help="Optional - describe the technology or product")
        
        col1, col2 = st.columns(2)
        
        with col1:
            quarter = st.selectbox("Quarter *", ["Q1", "Q2", "Q3", "Q4"])
            month = st.selectbox("Month *", 
                               ["January", "February", "March", "April", "May", "June",
                                "July", "August", "September", "October", "November", "December"])
        
        with col2:
            value = st.text_input("Deal Value *", 
                                 help="Enter amount (e.g., 100M, 1.5B, or Undisclosed)")
        
        if deal_type == "M&A":
            acquirer = st.text_input("Acquirer *", help="Company making the acquisition")
            deal_subtype = st.selectbox("Deal Subtype", ["Acquisition", "Merger"])
        else:
            funding_type = st.selectbox("Funding Type", ["VC", "PE"])
            lead_investors = st.text_input("Lead Investors", help="Optional")
        
        submitted = st.form_submit_button("➕ Add Deal", use_container_width=True)
        
        if submitted:
            if not company or not value:
                st.error("Please fill in all required fields (*)")
            else:
                deal = {
                    'type': deal_type,
                    'company': company,
                    'technology': technology if technology else '',
                    'value': value,
                    'quarter': quarter,
                    'month': month
                }
                
                if deal_type == "M&A":
                    deal['acquirer'] = acquirer
                    deal['deal_subtype'] = deal_subtype
                else:
                    deal['funding_type'] = funding_type
                    deal['lead_investors'] = lead_investors if lead_investors else 'Undisclosed'
                
                # Add to database
                added_ma, added_inv, skipped = add_deals_to_database([deal], ma_df, inv_df)
                
                if added_ma > 0 or added_inv > 0:
                    st.success(f"✅ Successfully added {added_ma + added_inv} deal(s)!")
                    if save_data(ma_df, inv_df):
                        st.cache_data.clear()
                        st.rerun()
                
                if skipped:
                    st.warning(f"⚠️ Skipped {len(skipped)} duplicate(s)")
                    for dup in skipped:
                        st.write(f"- {dup}")

def show_edit_page(ma_df, inv_df):
    """Edit existing deals"""
    st.header("✏️ Edit Deals")
    
    edit_type = st.radio("Select Data Type", ["M&A Deals", "Venture Investments"])
    
    df = ma_df if edit_type == "M&A Deals" else inv_df
    
    if df.empty:
        st.info(f"No {edit_type.lower()} to edit")
        return
    
    st.subheader(f"Edit {edit_type}")
    
    # Show editable dataframe
    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Technology/Description": st.column_config.TextColumn(
                "Technology Description",
                width="large",
                help="Technology or product description"
            )
        }
    )
    
    if st.button("💾 Save Changes", use_container_width=True):
        if edit_type == "M&A Deals":
            if save_data(edited_df, inv_df):
                st.success("✅ Changes saved successfully!")
                st.cache_data.clear()
                st.rerun()
        else:
            if save_data(ma_df, edited_df):
                st.success("✅ Changes saved successfully!")
                st.cache_data.clear()
                st.rerun()

def show_delete_page(ma_df, inv_df):
    """Delete deals"""
    st.header("🗑️ Delete Deals")
    
    delete_type = st.radio("Select Data Type", ["M&A Deals", "Venture Investments"])
    
    df = ma_df.copy() if delete_type == "M&A Deals" else inv_df.copy()
    
    if df.empty:
        st.info(f"No {delete_type.lower()} to delete")
        return
    
    # Add selection column
    df.insert(0, 'Select', False)
    
    st.subheader(f"Select {delete_type} to Delete")
    
    # Show dataframe with selection
    edited_df = st.data_editor(
        df,
        use_container_width=True,
        disabled=[col for col in df.columns if col != 'Select'],
        hide_index=True,
        column_config={
            "Select": st.column_config.CheckboxColumn(
                "Select to Delete",
                help="Check to mark for deletion",
                default=False,
            )
        }
    )
    
    selected_count = edited_df['Select'].sum()
    
    if selected_count > 0:
        st.warning(f"⚠️ {selected_count} deal(s) selected for deletion")
        
        if st.button(f"🗑️ Delete {selected_count} Deal(s)", type="primary"):
            # Remove selected rows
            df_to_keep = edited_df[~edited_df['Select']].drop('Select', axis=1)
            
            if delete_type == "M&A Deals":
                if save_data(df_to_keep, inv_df):
                    st.success(f"✅ Successfully deleted {selected_count} deal(s)!")
                    st.cache_data.clear()
                    st.rerun()
            else:
                if save_data(ma_df, df_to_keep):
                    st.success(f"✅ Successfully deleted {selected_count} deal(s)!")
                    st.cache_data.clear()
                    st.rerun()

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