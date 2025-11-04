import plotly.graph_objects as go
import pandas as pd
import streamlit as st
from utils.data_loader import format_currency, extract_numeric_value

def create_volume_chart(data, deal_type):
    """Create stacked bar chart with deal volume and line overlay for deal count"""
    
    # Determine the value column based on deal type
    if deal_type == "M&A":
        value_col = 'Deal Value'
    else:
        value_col = 'Amount Raised'
    
    # Filter out 'Undisclosed' quarters
    data_filtered = data[data['Quarter'] != 'Undisclosed'].copy()
    
    # Extract numeric values
    data_filtered['numeric_value'] = data_filtered[value_col].apply(extract_numeric_value)
    
    # Group by quarter
    quarterly_data = data_filtered.groupby('Quarter').agg({
        'numeric_value': 'sum',
        'Company': 'count'
    }).reset_index()
    
    quarterly_data.columns = ['Quarter', 'Total_Value', 'Deal_Count']
    
    # Sort quarters
    quarter_order = ['Q1', 'Q2', 'Q3', 'Q4']
    quarterly_data['Quarter'] = pd.Categorical(quarterly_data['Quarter'], categories=quarter_order, ordered=True)
    quarterly_data = quarterly_data.sort_values('Quarter')
    
    # Convert to billions
    quarterly_data['Total_Value_B'] = quarterly_data['Total_Value'] / 1_000_000_000
    
    # Create figure
    fig = go.Figure()
    
    # Add bar chart for deal value
    fig.add_trace(go.Bar(
        name='Deal Value ($B)',
        x=quarterly_data['Quarter'],
        y=quarterly_data['Total_Value_B'],
        marker_color='#1f77b4' if deal_type == "M&A" else '#2ca02c',
        text=[f"${x:.1f}B" for x in quarterly_data['Total_Value_B']],
        textposition='outside',
        yaxis='y'
    ))
    
    # Add line chart for deal count
    fig.add_trace(go.Scatter(
        name='Deal Count',
        x=quarterly_data['Quarter'],
        y=quarterly_data['Deal_Count'],
        mode='lines+markers+text',
        line=dict(color='#ff7f0e', width=3),
        marker=dict(size=10),
        text=[f"{int(x)} deals" for x in quarterly_data['Deal_Count']],
        textposition='top center',
        yaxis='y2'
    ))
    
    # Update layout
    fig.update_layout(
        height=500,
        xaxis=dict(title='Quarter', tickfont=dict(size=12)),
        yaxis=dict(
            title='Total Deal Value ($B)',
            side='left',
            titlefont=dict(size=12),
            tickfont=dict(size=11)
        ),
        yaxis2=dict(
            title='Number of Deals',
            overlaying='y',
            side='right',
            titlefont=dict(size=12),
            tickfont=dict(size=11)
        ),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        ),
        template='plotly_white',
        margin=dict(t=50, b=50, l=50, r=50)
    )
    
    return fig

def create_top_deals_card(data, deal_type):
    """Create cards for top 3 deals by value"""
    
    # Determine the value column based on deal type
    if deal_type == "M&A":
        value_col = 'Deal Value'
        company_col = 'Company'
        acquirer_col = 'Acquirer'
    else:
        value_col = 'Amount Raised'
        company_col = 'Company'
        acquirer_col = 'Lead Investors'
    
    # Filter out undisclosed values
    data_filtered = data[data[value_col] != 'Undisclosed'].copy()
    
    # Extract numeric values for sorting
    data_filtered['numeric_value'] = data_filtered[value_col].apply(extract_numeric_value)
    
    # Sort and get top 3
    top_deals = data_filtered.nlargest(3, 'numeric_value')
    
    if len(top_deals) == 0:
        st.info("No deals with disclosed values found")
        return
    
    # Display cards
    for idx, (_, deal) in enumerate(top_deals.iterrows(), 1):
        # Format the value
        formatted_value = format_currency(deal['numeric_value'])
        
        # Create the card
        with st.container():
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 1rem; color: #666; font-weight: 600;">#{idx}</div>
                <div class="deal-value">{formatted_value}</div>
                <div class="deal-company">{deal[company_col]}</div>
                <div style="font-size: 0.9rem; color: #666; margin-top: 0.5rem;">
                    {acquirer_col.replace('_', ' ')}: {deal[acquirer_col]}
                </div>
                <div style="font-size: 0.85rem; color: #888; margin-top: 0.5rem;">
                    {deal['Quarter']} • {deal['Month']}
                </div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

def create_jpm_summary_chart(jpm_data):
    """Create comprehensive JP Morgan summary chart"""
    
    fig = go.Figure()
    
    # Add grouped bars for each category
    categories = ['M&A', 'Venture', 'IPO', 'Licensing']
    colors = ['#1f77b4', '#2ca02c', '#ff7f0e', '#d62728']
    
    for i, category in enumerate(categories):
        amount_key = f"{category.lower()}_amount"
        
        fig.add_trace(go.Bar(
            name=f"{category} Amount ($B)",
            x=jpm_data['quarters'],
            y=jpm_data[amount_key],
            marker_color=colors[i],
            text=[f"${x}B" for x in jpm_data[amount_key]],
            textposition='outside'
        ))
    
    # Add volume line
    total_volume = [
        sum([jpm_data[f"{cat.lower()}_volume"][i] for cat in categories])
        for i in range(len(jpm_data['quarters']))
    ]
    
    fig.add_trace(go.Scatter(
        name='Total Deal Volume',
        x=jpm_data['quarters'],
        y=total_volume,
        mode='lines+markers+text',
        line=dict(color='#9467bd', width=3),
        marker=dict(size=10),
        text=[f"{x} deals" for x in total_volume],
        textposition='top center',
        yaxis='y2'
    ))
    
    fig.update_layout(
        barmode='group',
        height=600,
        xaxis=dict(title='Quarter'),
        yaxis=dict(title='Deal Amount ($B)', side='left'),
        yaxis2=dict(title='Deal Volume', overlaying='y', side='right'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        template='plotly_white'
    )
    
    return fig
