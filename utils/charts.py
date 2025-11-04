import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
from typing import Dict

def create_volume_chart(data_dict: Dict[str, pd.DataFrame]) -> go.Figure:
    """Create M&A volume comparison chart"""
    h1_count = len(data_dict['h1_ma'])
    h2_count = len(data_dict['h2_ma'])
    
    fig = go.Figure(data=[
        go.Bar(
            x=['H1 2024', 'H2 2024'],
            y=[h1_count, h2_count],
            text=[h1_count, h2_count],
            textposition='auto',
            marker_color=['#1f77b4', '#ff7f0e'],
            hovertemplate='<b>%{x}</b><br>Deals: %{y}<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title="M&A Activity Volume by Half",
        xaxis_title="Period",
        yaxis_title="Number of Deals",
        height=500,
        margin=dict(t=50, b=50, l=50, r=50),
        showlegend=False,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)'
    )
    
    return fig

def create_investment_chart(data_dict: Dict[str, pd.DataFrame]) -> go.Figure:
    """Create investment activity comparison chart"""
    h1_count = len(data_dict['h1_investment'])
    h2_count = len(data_dict['h2_investment'])
    
    fig = go.Figure(data=[
        go.Bar(
            x=['H1 2024', 'H2 2024'],
            y=[h1_count, h2_count],
            text=[h1_count, h2_count],
            textposition='auto',
            marker_color=['#2ca02c', '#d62728'],
            hovertemplate='<b>%{x}</b><br>Investments: %{y}<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title="Investment Activity Volume by Half",
        xaxis_title="Period",
        yaxis_title="Number of Investments",
        height=500,
        margin=dict(t=50, b=50, l=50, r=50),
        showlegend=False,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)'
    )
    
    return fig

def create_combined_chart(data_dict: Dict[str, pd.DataFrame]) -> go.Figure:
    """Create combined M&A and Investment activity chart"""
    periods = ['H1 2024', 'H2 2024']
    ma_counts = [len(data_dict['h1_ma']), len(data_dict['h2_ma'])]
    inv_counts = [len(data_dict['h1_investment']), len(data_dict['h2_investment'])]
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=periods,
        y=ma_counts,
        name='M&A Activity',
        marker_color='#1f77b4',
        text=ma_counts,
        textposition='auto'
    ))
    
    fig.add_trace(go.Bar(
        x=periods,
        y=inv_counts,
        name='Investment Activity',
        marker_color='#2ca02c',
        text=inv_counts,
        textposition='auto'
    ))
    
    fig.update_layout(
        title="Combined M&A and Investment Activity",
        xaxis_title="Period",
        yaxis_title="Number of Transactions",
        height=500,
        margin=dict(t=50, b=50, l=50, r=50),
        barmode='group',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    return fig

def create_pie_chart(data: pd.DataFrame, column: str, title: str) -> go.Figure:
    """Create a pie chart for categorical data"""
    if column not in data.columns:
        return go.Figure()
    
    value_counts = data[column].value_counts().head(10)
    
    fig = go.Figure(data=[go.Pie(
        labels=value_counts.index,
        values=value_counts.values,
        hole=0.3
    )])
    
    fig.update_layout(
        title=title,
        height=500,
        margin=dict(t=50, b=50, l=50, r=50)
    )
    
    return fig