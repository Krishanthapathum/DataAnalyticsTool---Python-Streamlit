import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Sales Dashboard", layout="wide")

# App Header
st.markdown("""
    <style>
        .fixed-header {
            position: fixed;
            top: 50px;
            left: 0;
            width: 100%;
            background: #008000;
            color: white;
            text-align: center;
            padding: 0px;
            z-index: 1000;
            border-bottom: 2px solid #2c5aa0;
            border-radius: 0 0 8px 8px;
            padding-left: 200px; /* Adjust based on header height */

        }
        .main > div {
            padding-top: 80px; /* Adjust based on header height */
        }
    </style>
    <div class="fixed-header">
        <h1>Sales Analytics Dashboard - Sales Company</h1>
    </div>
""", unsafe_allow_html=True)

import sys, os

base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
file_path = os.path.join(base_path, 'BA_Dataset.xlsx')
df = pd.read_excel(file_path, sheet_name='Orders')
                       
with st.sidebar:
    st.title("Filters")
    region_filter = st.multiselect("Select Region", df['Region'].unique(), default=df['Region'].unique())
    segment_filter = st.multiselect("Select Segment", df['Segment'].unique(), default=df['Segment'].unique())
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("Download CSV", csv, "filtered_data.csv", "text/csv")
    
    st.markdown("""
        <div style='margin-top: 50px;'>
            <small style='color: gray;'>Developed by</small><br>
            <b>22760-KGKP Premalal</b><br>
            <b>22708-KLD Nayanamini</b>
        </div>
        <div style='margin-top: 20px;'>
            <small style='color: gray;'>Module Lecture</small><br>
            <b>Mrs. Nethmi Weerasinghe</b><br>
        </div>
    """, unsafe_allow_html=True)
    
# Reset index after preprocessing
df.reset_index(drop=True, inplace=True)

# Apply filters
filtered_df = df[(df['Region'].isin(region_filter)) & (df['Segment'].isin(segment_filter))]

st.markdown("""<div style='margin-top: 30px;'></div>""", unsafe_allow_html=True)

# --------- Horizontal Tabs ---------
tab1, tab2 = st.tabs([" Visualizations", " Descriptive Statistics"])


with tab1:


    filtered_df = df[(df['Region'].isin(region_filter)) & (df['Segment'].isin(segment_filter))]

    st.markdown("""
        <div style='margin-top: 50px;'>

        </div>
    """, unsafe_allow_html=True)

    # KPIs
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown("""
            <div style="border: 1px solid gray; padding: 14px; border-radius: 8px;">
                💰 <b>Total Sales</b><br>
                <h3>Rs. {:,.2f}</h3>
            </div>
        """.format(filtered_df['Sales'].sum()), unsafe_allow_html=True)

    with col2:
        st.markdown("""
            <div style="border: 1px solid gray; padding: 14px; border-radius: 8px;">
                📈 <b>Total Profit</b><br>
                <h3>Rs. {:,.2f}</h3>
            </div>
        """.format(filtered_df['Profit'].sum()), unsafe_allow_html=True)

    with col3:
        st.markdown("""
            <div style="border: 1px solid gray; padding: 14px; border-radius: 8px;">
                🛒 <b>Orders</b><br>
                <h3>{}</h3>
            </div>
        """.format(filtered_df['Order ID'].nunique()), unsafe_allow_html=True)

    with col4:
        st.markdown("""
            <div style="border: 1px solid gray; padding: 14px; border-radius: 8px;">
                👤 <b>Unique Customers</b><br>
                <h3>{}</h3>
            </div>
        """.format(filtered_df['Customer ID'].nunique()), unsafe_allow_html=True)

    st.markdown("---")

    # ----------- Sales Charts ----------- #
    st.subheader("Sales Charts")
    col5, col6 = st.columns(2)

    # Sales by region
    with col5:
        sales_by_region = filtered_df.groupby('Region')['Sales'].sum()
        fig, ax = plt.subplots(facecolor='none')
        ax.set_facecolor('none')
        sales_by_region.plot(kind='bar', color='#4e79a7', ax=ax)
        ax.set_title("Sales by Region", color='white')
        ax.tick_params(axis='x', colors='white')  
        ax.tick_params(axis='y', colors='white')  
        ax.xaxis.label.set_color('white')
        ax.yaxis.label.set_color('white')  
        st.pyplot(fig)

    # Profit by Category
    with col6:
        profit_by_category = filtered_df.groupby('Category')['Profit'].sum()
        fig, ax = plt.subplots(facecolor='none')
        ax.set_facecolor('none')
        profit_by_category.plot(kind='bar', color='#59a14f', ax=ax)
        ax.set_title("Profit by Category", color='white')
        ax.tick_params(axis='x', colors='white') 
        ax.tick_params(axis='y', colors='white')  
        ax.xaxis.label.set_color('white')  
        ax.yaxis.label.set_color('white') 
        st.pyplot(fig)

    st.markdown("---")

    col7, col8 = st.columns(2)

    # Segment Distibution
    with col7:
        segment_distribution = filtered_df['Segment'].value_counts()
        fig, ax = plt.subplots(facecolor='none')
        ax.set_facecolor('none')
        wedges, texts, autotexts = ax.pie(segment_distribution, labels=segment_distribution.index, autopct='%1.1f%%', startangle=140)
        for text in texts + autotexts:
            text.set_color('white')
        ax.set_title("Segment Distribution", color='white')
        st.pyplot(fig)

    # Sales by Category
    with col8:
        category_sales = filtered_df.groupby('Category')['Sales'].sum()
        fig, ax = plt.subplots(facecolor='none')
        ax.set_facecolor('none')
        wedges, texts, autotexts = ax.pie(category_sales, labels=category_sales.index, autopct='%1.1f%%', startangle=140)
        for text in texts + autotexts:
            text.set_color('white')
        ax.set_title("Sales by Category", color='white')
        st.pyplot(fig)

    st.markdown("---")

    # Monthly Sales Trend
    monthly_sales = filtered_df.resample('M', on='Order Date')['Sales'].sum()
    fig, ax = plt.subplots(figsize=(12,4))
    ax.set_facecolor('none')  
    fig.patch.set_facecolor('none') 
    ax.plot(monthly_sales.index, monthly_sales.values, marker='o', color='blue')
    ax.set_title('Monthly Sales Trend', color='white')
    ax.tick_params(colors='white')
    st.pyplot(fig)

    st.markdown("---")

    # ----------- New Advanced Charts ----------- #
    st.subheader("Advanced Visualizations")

    st.markdown("### Correlation Matrix")
    corr = filtered_df[['Sales', 'Profit', 'Quantity', 'Discount']].corr()
    fig, ax = plt.subplots(figsize=(2, 2), facecolor='none')
    ax.set_facecolor('none')
    sns.heatmap(corr, annot=True, cmap="coolwarm", ax=ax, annot_kws={"size": 4}, cbar=False)
    ax.tick_params(colors='white', labelsize=4)
    ax.set_title('Correlation Matrix', color='white', fontsize=6)
    plt.tight_layout()
    st.pyplot(fig, use_container_width=False)

    st.markdown("---")

    #  Interactive Scatter Plot
    st.markdown("### Sales vs Profit Scatter Plot")
    fig = px.scatter(filtered_df, x='Sales', y='Profit', color='Category', size='Quantity', hover_data=['City'])
    st.plotly_chart(fig)

    st.markdown("---")

    #  Treemap
    st.markdown("### Sales Treemap (Region > Category > Sub-Category)")
    fig = px.treemap(filtered_df, path=['Region', 'Category', 'Sub-Category'], values='Sales', color='Profit', hover_data=['Sales'])
    st.plotly_chart(fig)

    st.markdown("---")

    #  Sunburst Chart
    st.markdown("### Sunburst Chart (Segment > Category > Sub-Category)")
    fig = px.sunburst(filtered_df, path=['Segment', 'Category', 'Sub-Category'], values='Sales', color='Sales')
    st.plotly_chart(fig)

    st.markdown("---")

    # 5. Funnel Chart
    st.markdown("### Sales Funnel")
    funnel_data = filtered_df.groupby('Ship Mode').agg({'Sales': 'sum'}).reset_index()
    fig = px.funnel(funnel_data, x='Sales', y='Ship Mode')
    st.plotly_chart(fig)

    st.markdown("---")

    # Cumulative Sales
    st.markdown("### Cumulative Sales Over Time")
    filtered_df = filtered_df.sort_values('Order Date')
    filtered_df['Cumulative Sales'] = filtered_df['Sales'].cumsum()
    fig = px.line(filtered_df, x='Order Date', y='Cumulative Sales', title='Cumulative Sales')
    st.plotly_chart(fig)

    st.markdown("---")

    # Sales target
    st.markdown("### Sales Target vs Actual")
    target = 3000000  # sample target
    actual = filtered_df['Sales'].sum()
    fig = go.Figure(go.Indicator(
        mode = "gauge+number+delta",
        value = actual,
        delta = {'reference': target},
        gauge = {'axis': {'range': [None, target*1.2]}, 'bar': {'color': "green"}},
        title = {'text': "Total Sales vs Target"}
    ))
    st.plotly_chart(fig)

    st.markdown("---")

    # 8. Sankey Diagram (Segment to Category Flow)
    st.markdown("### Simplified Sankey Diagram (Segment to Category Flow)")
    segments = list(filtered_df['Segment'].unique())
    categories = list(filtered_df['Category'].unique())
    all_nodes = segments + categories
    source = []
    target = []
    value = []
    for seg in segments:
        for cat in categories:
            sales = filtered_df[(filtered_df['Segment'] == seg) & (filtered_df['Category'] == cat)]['Sales'].sum()
            if sales > 0:
                source.append(all_nodes.index(seg))
                target.append(all_nodes.index(cat))
                value.append(sales)
    fig = go.Figure(data=[go.Sankey(
        node=dict(pad=15, thickness=20, label=all_nodes, color="lightblue"),
        link=dict(source=source, target=target, value=value)
    )])
    fig.update_layout(title_text="Segment to Category Flow", font_size=10)
    st.plotly_chart(fig)

with tab2:
        
    filtered_df = df[(df['Region'].isin(region_filter)) & (df['Segment'].isin(segment_filter))]

    st.markdown("""
        <div style='margin-top: 50px;'>

        </div>
    """, unsafe_allow_html=True)

    # KPIs
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown("""
            <div style="border: 1px solid gray; padding: 14px; border-radius: 8px;">
                💰 <b>Total Sales</b><br>
                <h3>Rs. {:,.2f}</h3>
            </div>
        """.format(filtered_df['Sales'].sum()), unsafe_allow_html=True)

    with col2:
        st.markdown("""
            <div style="border: 1px solid gray; padding: 14px; border-radius: 8px;">
                📈 <b>Total Profit</b><br>
                <h3>Rs. {:,.2f}</h3>
            </div>
        """.format(filtered_df['Profit'].sum()), unsafe_allow_html=True)

    with col3:
        st.markdown("""
            <div style="border: 1px solid gray; padding: 14px; border-radius: 8px;">
                🛒 <b>Orders</b><br>
                <h3>{}</h3>
            </div>
        """.format(filtered_df['Order ID'].nunique()), unsafe_allow_html=True)

    with col4:
        st.markdown("""
            <div style="border: 1px solid gray; padding: 14px; border-radius: 8px;">
                👤 <b>Unique Customers</b><br>
                <h3>{}</h3>
            </div>
        """.format(filtered_df['Customer ID'].nunique()), unsafe_allow_html=True)

    st.markdown("---")


    st.subheader(" Descriptive Statistics Summery")

    st.markdown("""
    <style>
        .calc-box {
            background-color: #1e1e1e;
            padding: 20px;
            margin-bottom: 25px;
            border: 1px solid #555;
            border-radius: 10px;
            color: #f0f0f0;
            font-family: monospace;
            text-align: center;
        }
        .calc-result {
            margin-top: 10px;
            font-size: 16px;
            font-weight: bold;
        }
    </style>
    """, unsafe_allow_html=True)

    # Mean
    mean_sales = filtered_df['Sales'].mean()
    total_sales = filtered_df['Sales'].sum()
    count_sales = filtered_df['Sales'].count()

    st.markdown(f"""<div class='calc-box'>""", unsafe_allow_html=True)
    st.latex(r"\mu = \frac{\sum_{i=1}^{N} x_i}{N}")
    st.markdown(f"<div class='calc-result'>Mean = {total_sales:,.2f} / {count_sales} = {mean_sales:,.2f}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Median
    median_sales = filtered_df['Sales'].median()
    st.markdown(f"""<div class='calc-box'>""", unsafe_allow_html=True)
    st.latex(r"\text{Median} = \text{Middle Value of Ordered Dataset}")
    st.markdown(f"<div class='calc-result'>Median = {median_sales:,.2f}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Mode
    mode_sales = filtered_df['Sales'].mode()[0]
    st.markdown(f"""<div class='calc-box'>""", unsafe_allow_html=True)
    st.latex(r"\text{Mode} = \text{Most Frequent Value}")
    st.markdown(f"<div class='calc-result'>Mode = {mode_sales:,.2f}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Range
    range_sales = filtered_df['Sales'].max() - filtered_df['Sales'].min()
    st.markdown(f"""<div class='calc-box'>""", unsafe_allow_html=True)
    st.latex(r"\text{Range} = \max(x) - \min(x)")
    st.markdown(f"<div class='calc-result'>Range = {filtered_df['Sales'].max():,.2f} - {filtered_df['Sales'].min():,.2f} = {range_sales:,.2f}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Variance
    variance_sales = filtered_df['Sales'].var()
    st.markdown(f"""<div class='calc-box'>""", unsafe_allow_html=True)
    st.latex(r"s^2 = \frac{\sum_{i=1}^{N} (x_i - \mu)^2}{N-1}")
    st.markdown(f"<div class='calc-result'>Variance = {variance_sales:,.2f}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Standard Deviation
    std_sales = filtered_df['Sales'].std()
    st.markdown(f"""<div class='calc-box'>""", unsafe_allow_html=True)
    st.latex(r"s = \sqrt{s^2} = \sqrt{\text{Variance}}")
    st.markdown(f"<div class='calc-result'>Std Dev = √{variance_sales:,.2f} = {std_sales:,.2f}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # IQR
    iqr_sales = filtered_df['Sales'].quantile(0.75) - filtered_df['Sales'].quantile(0.25)
    st.markdown(f"""<div class='calc-box'>""", unsafe_allow_html=True)
    st.latex(r"IQR = Q_3 - Q_1")
    st.markdown(f"<div class='calc-result'>IQR = {filtered_df['Sales'].quantile(0.75):,.2f} - {filtered_df['Sales'].quantile(0.25):,.2f} = {iqr_sales:,.2f}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Count / Min / Max
    st.markdown(f"""<div class='calc-box'>""", unsafe_allow_html=True)
    st.latex(r"\text{Additional Metrics}")
    st.markdown(f"""
    <div class='calc-result'>
    Sales Count = {filtered_df.shape[0]}<br>
    Min Value= {filtered_df['Sales'].min():,.2f}<br>
    Max Value= {filtered_df['Sales'].max():,.2f}
    </div>""", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
