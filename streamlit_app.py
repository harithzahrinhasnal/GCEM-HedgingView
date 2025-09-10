from collections import namedtuple
import altair as alt
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import datetime as dt
from datetime import datetime
from io import BytesIO
import io
import time
import xlsxwriter
from fpdf import FPDF
import base64
import math



# Define your discrete color sequence PETRONAS COLORS
color_discrete_sequence = [
    "#00b1a9",  # Original color - R000 G177 B169
    "#763f98",  # Original color - R118 G063 B152
    "#20419a",  # Original color - R032 G065 B154
    "#fdb924",  # Original color - R253 G185 B036
    "#bfd730",  # Original color - R191 G215 B048
    "#007b73",  # Shade of R000 G177 B169
    "#3a1d4c",  # Shade of R118 G063 B152
    "#101e4a",  # Shade of R032 G065 B154
    "#cc8b1c",  # Shade of R253 G185 B036
    "#8e9c1b",  # Shade of R191 G215 B048
    "#00524f",  # Darker shade of R000 G177 B169
    "#9a6cb3",  # Lighter shade of R118 G063 B152
    "#4367c5",  # Lighter shade of R032 G065 B154
    "#fcd05b",  # Lighter shade of R253 G185 B036
    "#d0e16a",  # Lighter shade of R191 G215 B048
    "#005f57"   # Even darker shade of R000 G177 B169
]

color_discrete_sequence_2 = [
    '#00B1A9', '#763F98', '#FF6F61', '#3B5998', '#FFD166', 
    '#06D6A0', '#118AB2', '#073B4C', '#A05195', '#2EC4B6', 
    '#FFB703', '#E63946', '#457B9D', '#E9C46A', '#2A9D8F', 
    '#8D99AE', '#EF476F'
]



# Setting Up
st.set_page_config(page_title = "DashBoard",page_icon = r'Resources/4953098.png',layout ="wide")

st.markdown(
    """
        <style>
            .appview-container .main .block-container {{
                padding-top: {padding_top}rem;
                padding-bottom: {padding_bottom}rem;
                }}

        </style>""".format(
        padding_top=1, padding_bottom=1
    ),
    unsafe_allow_html=True,
)


df = pd.read_excel("PCHP Data.xlsx","Overall_data")


# Sidebar header and widgets for selecting filters
st.sidebar.header("Choose your filter:")
all_counterparties = df["FO.CounterpartyName"].dropna().unique()
all_portfolios = df["Portfolio"].dropna().unique()
all_dates = pd.to_datetime(df['FO.TradeDate']).dt.date.dropna().unique()
all_dealers = df["FO.DealerID"].dropna().unique()

# Add "All" option to the lists
all_counterparties = ['All'] + all_counterparties.tolist()
all_portfolios = ['All'] + all_portfolios.tolist()
all_dealers = ['All'] + all_dealers.tolist()

# Set default selections to include "All"
selected_counterparties = st.sidebar.multiselect("Counterparty", all_counterparties, default=['All'])
selected_portfolios = st.sidebar.multiselect("Portfolio", all_portfolios, default=['FY2026 PCHP'])

selected_dealers = st.sidebar.multiselect("Dealer", all_dealers, default=['All'])

# Update the selected options if "All" is selected
if 'All' in selected_counterparties:
    selected_counterparties = all_counterparties[1:]  # Exclude "All"
else:
    selected_counterparties = selected_counterparties

if 'All' in selected_portfolios:
    selected_portfolios = all_portfolios[1:]  # Exclude "All"
else:
    selected_portfolios = selected_portfolios

if 'All' in selected_dealers:
    selected_dealers = all_dealers[1:]  # Exclude "All"
else:
    selected_dealers = selected_dealers

# Filter data based on selected counterparties, portfolios, and date range
filtered_df = df[(df['FO.CounterpartyName'].isin(selected_counterparties)) &
                  (df['Portfolio'].isin(selected_portfolios))]

filtered_df = filtered_df[(filtered_df['FO.DealerID'].isin(selected_dealers))] 
                  

# Convert the "FO.TradeDate" column to datetime if it's not already
filtered_df['FO.TradeDate'] = pd.to_datetime(filtered_df['FO.TradeDate'], errors='coerce')



# Date range selection
st.sidebar.header("Select Date Range")

 ## Range selector
format = 'MMM DD, YYYY'  # format output

# Handle NaTType error and set default values for the date range
try:
    MIN_MAX_RANGE = (filtered_df['FO.TradeDate'].dropna().min(), filtered_df['FO.TradeDate'].dropna().max())
except KeyError:
    # Handle KeyError (e.g., due to NaTType) by setting default min and max dates
    MIN_MAX_RANGE = (pd.Timestamp('1900-01-01'), pd.Timestamp('2100-12-31'))

# Get the minimum and maximum dates from the filtered DataFrame
min_date = MIN_MAX_RANGE[0]
max_date = MIN_MAX_RANGE[1]

# Set the pre-selected dates to match the minimum and maximum dates
PRE_SELECTED_DATES = (min_date.to_pydatetime(), max_date.to_pydatetime())  # Convert to datetime objects

# Handle the KeyError (NaTType) when creating the slider
try:
    selected_min, selected_max = st.sidebar.slider(
        "Datetime slider",
        value=PRE_SELECTED_DATES,
        min_value=MIN_MAX_RANGE[0],
        max_value=MIN_MAX_RANGE[1],format=format
    )
except KeyError:
    # Set default values for the slider in case of NaTType error
    selected_min, selected_max = PRE_SELECTED_DATES

# Convert the date range to pandas Timestamp objects
start_date = pd.to_datetime(selected_min)
end_date = pd.to_datetime(selected_max)

filtered_df = filtered_df[(filtered_df['FO.TradeDate'] >= start_date) &
                  (filtered_df['FO.TradeDate'] <= end_date)]

ITM_df=filtered_df.copy()

#Title

st.title("Group Commodity Exposure Management Dashboard")
tab1, tab2, tab3,tabITM, tab4 = st.tabs(["Overall Data", "Overview", "MTM","ITM", "Report"])


with tab1:
    # Display PCHP Data
    st.title("PCHP Execution Data")
    # Create a formatted copy of the filtered DataFrame to preserve the original data
    formatted_df = filtered_df.copy()

    # Format date columns for better readability
    date_columns = ['FO.TradeDate', 'FO.StartFixDate', 'FO.EndFixDate', 'FO.Settlement_DeliveryDate']
    for column in date_columns:
        formatted_df[column] = formatted_df[column].dt.strftime('%d %b %Y')

    # Specify columns to display in the table
    columns_to_display = ['FO.TradeDate','FO.DealerID',"Portfolio", 'FO.CounterpartyName','FO.PremiumStrike1','FO.PremiumStrike2','FO.NetPremium', 'FO.Position_Quantity',
                        'FO.StrikePrice1', 'FO.StrikePrice2', 'FO.StartFixDate', 'FO.EndFixDate', 'FO.Settlement_DeliveryDate',
                        'O.January','O.February','O.March','O.April','O.May','O.June','O.July',
                        'O.August','O.September','O.October','O.November','O.December']
    
    

    # Reset index to start from 1
    formatted_df = formatted_df.reset_index(drop=True)

    # Start index from 1
    formatted_df.index = formatted_df.index + 1

    # Show the formatted DataFrame using st.dataframe
    st.dataframe(formatted_df[columns_to_display],height=500, use_container_width = True)

with tab2:
    st.title('Execution Overview')
    #Brent Price Data
    df_prices = pd.read_excel("PCHP Data.xlsx","Brent_Prices")
    df_prices['Date'] = pd.to_datetime(df_prices['Date'], errors='coerce')
    filtered_df_prices = df_prices[(df_prices['Date'] >= start_date) &
                    (df_prices['Date'] <= end_date)]

    # Create a Plotly line chart
    fig_Brent = px.line(filtered_df_prices, x='Date', y='Historical Brent Price', title='Trade Execution Window',
                        labels={'Historical Brent Price': 'Brent Price'})

    # Update the line color
    fig_Brent.update_traces(line_color='#808080')

    # Use the custom color sequence for each portfolio
    portfolio_colors = {portfolio: color_discrete_sequence[i % len(color_discrete_sequence)]
                        for i, portfolio in enumerate(filtered_df['Portfolio'].unique())}

    # Create a dictionary to store traces for each portfolio
    portfolio_traces = {}

    # Add markers for executed trades, grouped by portfolio
    for index, trade in filtered_df.iterrows():
        portfolio = trade['Portfolio']
        if portfolio not in portfolio_traces:
            portfolio_traces[portfolio] = {'x': [], 'y': [], 'color': portfolio_colors.get(portfolio, 'red')}
        
        corresponding_price = filtered_df_prices.loc[filtered_df_prices['Date'] == trade['FO.TradeDate'], 'Historical Brent Price'].values
        if len(corresponding_price) > 0:
            portfolio_traces[portfolio]['x'].append(trade['FO.TradeDate'])
            portfolio_traces[portfolio]['y'].append(corresponding_price[0])

    # Add the portfolio traces to the figure
    for portfolio, trace_data in portfolio_traces.items():
        fig_Brent.add_trace(
            go.Scatter(
                x=trace_data['x'],
                y=trace_data['y'],
                mode='markers',
                marker=dict(color=trace_data['color']),
                name=f'Executed Trades - {portfolio}',
                legendgroup=portfolio,
                showlegend=True
            )
        )

    # Update the layout to include the markers and show legend below x-axis
    fig_Brent.update_layout(
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )

     # Convert the chart to an image with higher resolution
    image = fig_Brent.to_image(format="png", width=1200, height=500, scale=2.0)

    # Save the image to a file
    image_path = r"Resources\Plots\Brent.png"
    with open(image_path, "wb") as f:
        f.write(image)

    # Display the Plotly chart
    st.plotly_chart(fig_Brent, use_container_width=True, height=400,key='1')



    # Calculate Total_Position_Quantity and Weighted_Avg_Net_Premium for each portfolio
    filtered_df['Weighted_Avg_Net_Premium'] = (filtered_df['FO.NetPremium'] * filtered_df['FO.Position_Quantity']) / filtered_df.groupby('Portfolio')['FO.Position_Quantity'].transform('sum')
    filtered_df['Weighted_Avg_Protection'] = (filtered_df['FO.StrikePrice1'] * filtered_df['FO.Position_Quantity']) / filtered_df.groupby('Portfolio')['FO.Position_Quantity'].transform('sum')
    filtered_df['Weighted_Avg_Lower_Protection'] = (filtered_df['FO.StrikePrice2'] * filtered_df['FO.Position_Quantity']) / filtered_df.groupby('Portfolio')['FO.Position_Quantity'].transform('sum')
    filtered_df['Weighted_Avg_Protection_Band'] = filtered_df['Weighted_Avg_Protection'] - filtered_df['Weighted_Avg_Lower_Protection']
    filtered_df['Total_Cost'] = (filtered_df['FO.NetPremium'] * filtered_df['FO.Position_Quantity'])
    grouped_data = filtered_df.groupby('Portfolio').agg(
    Total_Position_Quantity=pd.NamedAgg(column='FO.Position_Quantity', aggfunc='sum'),
    Trade_Numbers=pd.NamedAgg(column='Portfolio', aggfunc='count'),
    Total_Cost=pd.NamedAgg(column='Total_Cost', aggfunc='sum'),
    Weighted_Avg_Net_Premium=pd.NamedAgg(column='Weighted_Avg_Net_Premium', aggfunc='sum'),
    Weighted_Avg_Protection=pd.NamedAgg(column='Weighted_Avg_Protection', aggfunc='sum'),
    Weighted_Avg_Lower_Protection=pd.NamedAgg(column='Weighted_Avg_Lower_Protection', aggfunc='sum'),
    Weighted_Avg_Protection_Band=pd.NamedAgg(column='Weighted_Avg_Protection_Band', aggfunc='sum')
    ).reset_index()

    

    # Apply accounting format to the numeric columns
    grouped_data['Total_Position_Quantity'] = grouped_data['Total_Position_Quantity'].apply('{:,.0f}'.format)
    grouped_data['Total_Cost'] = grouped_data['Total_Cost'].apply('USD {:,.2f}'.format)
    grouped_data['Weighted_Avg_Net_Premium'] = grouped_data['Weighted_Avg_Net_Premium'].apply('USD {:,.2f}'.format)
    grouped_data['Weighted_Avg_Protection'] = grouped_data['Weighted_Avg_Protection'].apply('USD {:,.2f}'.format)
    grouped_data['Weighted_Avg_Lower_Protection'] = grouped_data['Weighted_Avg_Lower_Protection'].apply('USD {:,.2f}'.format)
    grouped_data['Weighted_Avg_Protection_Band'] = grouped_data['Weighted_Avg_Protection_Band'].apply('USD {:,.2f}'.format)


    


   # Define the color for all cells
    grey_color = '#f6f6f6'

    fig3 = go.Figure(data=[go.Table(
        header=dict(
            values=['Portfolio', 'Number of Trades', 'Total Volume Hedged', 'Total Cost', 'Weighted Average Net Premium', 'Weighted Average Protection', 'Weighted Average Lower Protection', 'Protection Band'],
            # Applying colors to the header cells
            line_color='black', align='center'
        ),
        cells=dict(
            values=[grouped_data['Portfolio'], grouped_data['Trade_Numbers'], grouped_data['Total_Position_Quantity'], grouped_data['Total_Cost'], grouped_data['Weighted_Avg_Net_Premium'],
                    grouped_data['Weighted_Avg_Protection'], grouped_data['Weighted_Avg_Lower_Protection'], grouped_data['Weighted_Avg_Protection_Band']],
            # Applying colors to all cells
             # Setting the border color
            line_color='black', align='center'
        )
    )])

    
    

        # Grouped data with more descriptive column names
    grouped_data = grouped_data.rename(columns={
        'Portfolio': 'Portfolio',
        'Total_Position_Quantity': 'Total Position Quantity',
        'Trade_Numbers': 'Number of Trades',
        'Total_Cost': 'Total Cost',
        'Weighted_Avg_Net_Premium': 'Weighted Average Net Premium',
        'Weighted_Avg_Protection': 'Weighted Average Protection',
        'Weighted_Avg_Lower_Protection': 'Weighted Average Lower Protection',
        'Weighted_Avg_Protection_Band': 'Weighted Average Protection Band'
    })

    # Apply text-align: center to all cells in the DataFrame
    grouped_data = grouped_data.style.set_properties(**{'text-align': 'center'})

    # Display the grouped data with center-aligned values using st.dataframe()
    #st.dataframe(grouped_data, use_container_width=True, hide_index=True)
    
    st.plotly_chart(fig3, use_container_width=True,key="2")

     # Convert the chart to an image with higher resolution
    image = fig3.to_image(format="png", width=1200, height=250, scale=2.0)

    # Save the image to a file
    image_path = r"Resources\Plots\Execution_table.png"
    with open(image_path, "wb") as f:
        f.write(image)


    st.divider()

    col1, col2 = st.columns((2))

    with col1:
        # Custom colors for each dealer
        dealer_colors = {
            'HZ': '#00b1a9',
            'DS': '#763f98',
            'EG': "#20419a",
            'AS': "#fdb924",
            # Add more dealers and colors as needed
        }    
        
        # Calculate Volume executed versus Counterparty
        st.subheader("Volume executed versus Counterparty")
        
        # Add a column for custom colors based on DealerID
        filtered_df['Color'] = filtered_df['FO.DealerID'].map(dealer_colors)

        # Create the histogram with custom colors
        fig1 = px.histogram(filtered_df, x='FO.Acronym', y='FO.Position_Quantity', color='FO.DealerID', title='Sum of Volume Executed', color_discrete_map=dealer_colors)

        # Update the x-axis category order
        fig1.update_xaxes(categoryorder='total descending')

        # Rename x and y labels
        fig1.update_xaxes(title_text='Counterparties')
        fig1.update_yaxes(title_text='Quantity, bbls')

        # Add values at the top of each bar
        fig1.update_traces(texttemplate='%{y}', textposition='inside')

        # Show the Plotly figure in Streamlit
        st.plotly_chart(fig1, use_container_width=True, height=200)

        # Convert the chart to an image
        image = fig1.to_image(format="png", width=1200, height=550, scale=2.0)

        # Save the image to a file
        image_path = r"Resources\Plots\volume_dealer.png"
        with open(image_path, "wb") as f:
            f.write(image)



    with col1:
        df_refresh = pd.read_excel("PCHP Data.xlsx", "Sheet_Info")

        # Assuming 'Date_today' contains a single date value in the DataFrame
        date_today_value = df_refresh['Date_today'].iloc[0]

        # Convert to a datetime object and format it
        date_limit = datetime.strptime(str(date_today_value), "%Y-%m-%d %H:%M:%S.%f")

        # Format the datetime object to the desired format
        formatted_date_limit = date_limit.strftime("%d %b %Y")

        df_limits = pd.read_excel("PCHP Data.xlsx","Credit_Limit_data")
        st.subheader("Available Limits")
        df_limits = pd.read_excel("PCHP Data.xlsx","Credit_Limit_data")
        fig_limits = px.bar(df_limits, x='FO.Acronym', y=['Available Volume Limit', 'Volume Utilised'],
                title='Volume Limit and Volume Utilized by Counterparty as of '+ formatted_date_limit,color_discrete_sequence=color_discrete_sequence)
        
        # Rename x and y labels
        fig_limits.update_yaxes(title_text='Quantity, bbls')
        # Add values at the top of each bar
        fig_limits.update_traces(texttemplate='%{y}', textposition='inside')
        #fig_limits .update_xaxes(categoryorder='total descending')
        st.plotly_chart(fig_limits, use_container_width=True, height=200)       

    with col2:
        st.subheader("Monthly Volume Executed")
        # Reshape the data to have 'Month' as a column and corresponding values
        df_melted = pd.melt(filtered_df, id_vars=['Portfolio'], value_vars=['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                            var_name='Month', value_name='Value')

        # Define the correct order of months
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

        # Convert 'Month' to a categorical data type with the correct order
        df_melted['Month'] = pd.Categorical(df_melted['Month'], categories=month_order, ordered=True)

        # Group by Portfolio, Month, and Value type (Quantity or Premium) and sum the values
        df_grouped = df_melted.groupby(['Portfolio', 'Month']).sum().reset_index()
    
        # Create a line chart for quantities
        fig_quantity = px.bar(df_grouped, x='Month', y='Value', color='Portfolio',
                            title='Quantity Comparison by Portfolio for Each Month',
                            labels={'Value': 'Quantity'}, barmode='group')



        # Add a horizontal line to indicate the targeted value
        default_targeted_value =  int(232,680,000 /12)  # Adjust this value according to your targeted value
        targeted_value = [default_targeted_value,default_targeted_value,default_targeted_value,
                          default_targeted_value,default_targeted_value,default_targeted_value,
                          default_targeted_value,default_targeted_value,default_targeted_value,
                          default_targeted_value,default_targeted_value,default_targeted_value]

        # Create a trace for the target line
        target_trace = go.Scatter(x=df_grouped['Month'], y=[targeted_value],
                                mode='lines', line=dict(color='orange', dash='dash'),
                                name='FY2026 Mandated Volume')

        # Add the target trace to the figure
        fig_quantity.add_trace(target_trace)
        if len(selected_portfolios) == 1:
            # Calculate unexecuted volumes by subtracting executed volumes from the targeted value
            df_grouped['Unexecuted'] = targeted_value - df_grouped['Value']

            if selected_portfolios == ['FY2026 PCHP']:
                # Create a stacked bar chart with custom colors
                fig_stacked_bar = px.bar(df_grouped, x='Month', y=['Value', 'Unexecuted'],
                                        title='Executed vs. Unexecuted Volumes by Portfolio for Each Month',
                                        labels={'Value': 'Executed', 'Unexecuted': 'Unexecuted'},
                                        barmode='stack',color_discrete_sequence=color_discrete_sequence)
            else:
                # Create a stacked bar chart with custom colors
                fig_stacked_bar = px.bar(df_grouped, x='Month', y=['Value'], color='Portfolio',
                                        title='Executed vs. Unexecuted Volumes by Portfolio for Each Month',
                                        barmode='stack',color_discrete_sequence=color_discrete_sequence)


            # Rename x and y labels
            fig_stacked_bar.update_yaxes(title_text='Quantity, bbls')
            # Add values at the top of each bar
            fig_stacked_bar.update_traces(texttemplate='%{y}', textposition='inside')
            # Set the color for "Unexecuted" bars to red
            fig_stacked_bar.update_traces(marker_color='red', selector=dict(name='Unexecuted'))

            st.plotly_chart(fig_stacked_bar, use_container_width=True, height=200)

             # Convert the chart to an image
            image = fig_stacked_bar.to_image(format="png", width=1200, height=350, scale=2.0)

            # Save the image to a file
            image_path = r"Resources\Plots\volume_executed.png"
            with open(image_path, "wb") as f:
                f.write(image)

            
        else:
            st.write("No data available for visualization.")

    with col2:
        st.subheader("Counterparty Monthly Volume Executed")
        # Reshape the data to have 'Month' as a column and corresponding values
        df_melted = pd.melt(filtered_df, id_vars=['FO.Acronym'], value_vars=['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                            var_name='Month', value_name='Value')

        # Define the correct order of months
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

        # Convert 'Month' to a categorical data type with the correct order
        df_melted['Month'] = pd.Categorical(df_melted['Month'], categories=month_order, ordered=True)

        # Group by Portfolio, Month, and Value type (Quantity or Premium) and sum the values
        df_grouped = df_melted.groupby(['FO.Acronym', 'Month']).sum().reset_index()

        fig_quantity = go.Figure()

        for i, counterparty in enumerate(df_grouped['FO.Acronym'].unique()):
            df_counterparty = df_grouped[df_grouped['FO.Acronym'] == counterparty]
            fig_quantity.add_trace(go.Bar(
                x=df_counterparty['Month'],
                y=df_counterparty['Value'],
                name=counterparty,
                marker_color=color_discrete_sequence[i % len(color_discrete_sequence_2)],
                text=df_counterparty['Value'],  # Use y-values as text
                textposition='inside',
                texttemplate='%{text:.2s}',
            ))

        fig_quantity.update_layout(
            title='Quantity Comparison by Counterparty for Each Month',
            xaxis_title='Month',
            yaxis_title='Quantity',
            barmode='stack' 
        )
        
        
        st.plotly_chart(fig_quantity, use_container_width=True, height=200)

         # Convert the chart to an image
        image = fig_quantity.to_image(format="png", width=1200, height=600, scale=2.0)

        # Save the image to a file
        image_path = r"Resources\Plots\volume_cp.png"
        with open(image_path, "wb") as f:
            f.write(image)


# Assuming selected_portfolio is a list
if len(selected_portfolios) > 0 and 'All' not in selected_portfolios:
    # Ensure only one element in selected_portfolio
    selected_portfolio = [selected_portfolios[0]]

def visualize_data(st, filtered_df, strike_price_column, strike_price_name):
    if not filtered_df.empty:
        # Remove rows with NaN values in the "FO.TransactionNumber" column
        filtered_df = filtered_df.dropna(subset=['FO.TransactionNumber'])

        # Remove rows with NaN values in the "Total Outstanding" column
        filtered_df = filtered_df.dropna(subset=['Total Outstanding'])

        # Check if "Total Outstanding" column is not empty
        if not filtered_df['Total Outstanding'].empty:
            # Remove rows with 0 values in the "Total Outstanding" column
            filtered_df = filtered_df[filtered_df['Total Outstanding'] != 0]

            # Extract relevant columns for visualization
            months = ['O.January','O.February','O.March','O.April','O.May','O.June','O.July',
                        'O.August','O.September','O.October','O.November','O.December']
            monthly_data = filtered_df[months]

            # Group by strike_price_column and sum the data
            grouped_data = filtered_df.groupby(strike_price_column)[months].sum()

            # Check if grouped_data is not empty
            if not grouped_data.empty:
                # Transpose the data for plotting
                transposed_data = grouped_data.transpose()

                # Plotting the data using Plotly
                fig = go.Figure()



                # Add bar trace for each Strike Price
                for i, strike_price in enumerate(transposed_data.columns):
                    fig.add_trace(go.Bar(
                        x=transposed_data.index,
                        y=transposed_data[strike_price],
                        name=f'Strike Price {strike_price}',
                        marker_color=color_discrete_sequence[i % len(color_discrete_sequence)],
                        text=transposed_data[strike_price],  # Use y-values as text
                        textposition='outside',
                        texttemplate='%{text:.2s}',
                    ))

                fig.update_layout(
                    xaxis_title="Months",
                    yaxis_title="Total Outstanding Barrels, bbls",
                    title='Total Volume according to Strike Levels',
                    xaxis_tickangle=-45,
                    barmode='stack',
                    showlegend=True,
                    legend=dict(title=strike_price_name, x=1, y=1),
                    
                )
                # Add values at the top of each bar
                fig.update_traces(texttemplate='%{y:.2s}', textposition='inside')

                #custom_tick_labels = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

                # Set custom tick values and labels for the x-axis
                #fig.update_xaxes(tickvals=transposed_data.index, ticktext=custom_tick_labels)


                # Display the Plotly chart
                st.plotly_chart(fig, use_container_width=True)
                

                # Display table
                #st.subheader("Volume Breakdown")
                #st.dataframe(grouped_data,height=150, use_container_width = True)

                # Convert the chart to an image
                image = fig.to_image(format="png", width=1200, height=600, scale=2.0)

                # Save the image to a file
                image_path = rf"Resources\Plots\Outstanding_{strike_price_name}.png"
                with open(image_path, "wb") as f:
                    f.write(image)

            else:
                st.write("No data available for visualization.")
        else:
            st.write("No data available for visualization. Total Outstanding column is empty.")
    else:
        st.write("No data available for visualization.")


def strike_data(st, filtered_df, strike_price_column, strike_price_name):
    if not filtered_df.empty:
        # Remove rows with NaN values in the "FO.TransactionNumber" column
        filtered_df = filtered_df.dropna(subset=['FO.TransactionNumber'])

        # Remove rows with NaN values in the "Total Outstanding" column
        filtered_df = filtered_df.dropna(subset=['Total Outstanding'])

        # Check if "Total Outstanding" column is not empty
        if not filtered_df['Total Outstanding'].empty:
            # Remove rows with 0 values in the "Total Outstanding" column
            filtered_df = filtered_df[filtered_df['Total Outstanding'] != 0]

            # Extract relevant columns for visualization
            months = ['O.January','O.February','O.March','O.April','O.May','O.June','O.July',
                        'O.August','O.September','O.October','O.November','O.December']
            monthly_data = filtered_df[months]

            # Group by strike_price_column and sum the data
            grouped_data = filtered_df.groupby(strike_price_column)[months].sum()

            # Check if grouped_data is not empty
            if not grouped_data.empty:
                # Transpose the data for plotting
                transposed_data = grouped_data.transpose()

                return grouped_data
                
with tab3:
    st.title("Mark to Market Data")
    col1, col2 = st.columns((2))
    
    # Check if "Total Outstanding" column is not empty
    if not filtered_df['Total Outstanding'].empty:
        with col1:
            st.subheader("Outstanding Upper Strike Level")
            visualize_data(st, filtered_df, 'FO.StrikePrice1', 'FO.StrikePrice1')
        with col2:
            st.subheader("Outstanding Lower Strike Level")
            visualize_data(st, filtered_df, 'FO.StrikePrice2', 'FO.StrikePrice2')
    else:
        st.write("No data available for visualization.")

    st.divider()
    st.title("BBG Option Price and Valuation")
    try:
        # Read the Excel file
        df_BBG = pd.read_excel("BBG_Output.xlsx", sheet_name=None)

        # Get all sheet names
        sheet_names = list(df_BBG.keys())

        # Create a dropdown to select sheet
        default_sheet = sheet_names[-1]  # Set the default value to the last sheet name
        selected_sheet = st.selectbox("Select a sheet", sheet_names, index=len(sheet_names)-1)

        # Show the selected sheet data
        st.write("Data Refreshed:", selected_sheet)

        # Rename the first column
        df_selected_sheet = df_BBG[selected_sheet].rename(columns={df_BBG[selected_sheet].columns[0]: 'Strike Price'})

        # Convert numerical values in the first column (except the last one) to integers with one decimal place
        for i in range(len(df_selected_sheet) - 1):
            value = df_selected_sheet.iloc[i, 0]
            if isinstance(value, (int, float)):
                df_selected_sheet.iloc[i, 0] = round(float(value), 1)

        # Convert the rounded numerical values to integers
        df_selected_sheet.iloc[:-1, 0] = df_selected_sheet.iloc[:-1, 0].astype(int)

        st.dataframe(df_selected_sheet, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Error: {e}")

    st.divider()

    def process_dataframe(df1, df2):
        # Find common values in the first column of both dataframes
        common_values = df1.iloc[:, 0].isin(df2.iloc[:, 0])
        
        # Filter df1 and df2 based on common values in the first column
        df1_filtered = df1[df1.iloc[:, 0].isin(df2.iloc[:, 0])]
        df2_filtered = df2[df2.iloc[:, 0].isin(df1.iloc[:, 0])]
        
        df1_filtered = df1_filtered.reset_index(drop=True)
        df2_filtered = df2_filtered.reset_index(drop=True)
        # Reindex df2 to match the row and column indices of df1
        df2_reindexed = df2_filtered.reindex(index=df1_filtered.index, columns=df1_filtered.columns)
       
        # Multiply corresponding elements from df1 and df2
        result_df = df1_filtered * df2_reindexed

        # Assign the first column from df2 to the corresponding column in the result_df
        result_df[df1_filtered.columns[0]] = df1_filtered[df1_filtered.columns[0]]

        # Replace NaN values in result_df with 0
        result_df.fillna(0, inplace=True)

        return result_df

    if selected_portfolios == ['FY2025 PCHP']:
        # Process the first set of data
        df1 = df_selected_sheet
        df2 = strike_data(st, filtered_df, 'FO.StrikePrice1', 'FO.StrikePrice1')
        df2.reset_index(inplace=True)
        for column in df2.columns:
            df2[column] = df2[column].astype(int)
        df_Upper = process_dataframe(df1, df2)

        # Process the second set of data
        df1 = df_selected_sheet
        df3 = strike_data(st, filtered_df, 'FO.StrikePrice2', 'FO.StrikePrice2')
        df3.reset_index(inplace=True)
        for column in df3.columns:
            df3[column] = df3[column].astype(int)
        df_Lower = process_dataframe(df1, df3)

        # Display the results
        col3, col4 = st.columns((2))
    
        with col3:
            
            # Transpose the DataFrame to have months as columns and Strike Price as index
            df_Upper_transposed = df_Upper.set_index('Strike Price').transpose()

            # Create a Plotly bar chart
            fig = go.Figure()

            # Add bar trace for each Strike Price
            for i, strike_price in enumerate(df_Upper_transposed.columns):
                fig.add_trace(go.Bar(
                    x=df_Upper_transposed.index,
                    y=df_Upper_transposed[strike_price],
                    name=f'Strike Price {strike_price}',
                    marker_color=color_discrete_sequence[i % len(color_discrete_sequence)],
                    text=df_Upper_transposed[strike_price],  # Use y-values as text
                    textposition='outside',
                    showlegend=True,
                    texttemplate='%{text:.2s}',
                ))

            # Update layout with axis labels and title
            fig.update_layout(xaxis_title='Tenure',
                            yaxis_title='Value, USD',
                            title='Valuation of Upper Put Options',legend=dict(x=0, y=1.0))
            
            # Add values at the top of each bar
            fig.update_traces(texttemplate='%{y:.2s}', textposition='outside')

            custom_tick_labels = ['January Outstanding', 'February Outstanding', 'March Outstanding', 'April Outstanding', 'May Outstanding', 'June Outstanding', 'July Outstanding', 'August Outstanding', 'September Outstanding', 'October Outstanding', 'November Outstanding', 'December Outstanding']


            # Show plot
            st.plotly_chart(fig)
            # Print DataFrame
            st.dataframe(df_Upper, height=150, use_container_width=True, hide_index=True)

            # Convert the chart to an image
            image = fig.to_image(format="png", width=1200, height=550, scale=2.0)

            # Save the image to a file
            image_path = r"Resources\Plots\upper_put_options.png"
            with open(image_path, "wb") as f:
                f.write(image)

            # Set up the file name
            filename = "plotly_chart.png"
            # Convert the image to bytes
            image_bytes_2 = io.BytesIO(image)
            # Trigger the download
            st.download_button(label="Download Image", data=image_bytes_2, file_name=filename, mime="image/png", key="download_button_1")
        
        with col4:
            # Transpose the DataFrame to have months as columns and Strike Price as index
            df_Lower_transposed = df_Lower.set_index('Strike Price').transpose()

            # Create a Plotly bar chart
            fig2 = go.Figure()


            # Add bar trace for each Strike Price
            for i, strike_price in enumerate(df_Lower_transposed.columns):
                fig2.add_trace(go.Bar(
                    x=df_Lower_transposed.index,
                    y=df_Lower_transposed[strike_price],
                    name=f'Strike Price {strike_price}',
                    marker_color=color_discrete_sequence[i % len(color_discrete_sequence)],
                    text=df_Lower_transposed[strike_price],  # Use y-values as text
                    textposition='outside',showlegend=True,
                    texttemplate='%{text:.2s}',
                ))

            # Update layout with axis labels and title
            fig2.update_layout(xaxis_title='Tenure',
                            yaxis_title='Value, USD',
                            title='Valuation of Lower Put Options',legend=dict(x=0, y=1.0))
            
            custom_tick_labels = ['January Outstanding', 'February Outstanding', 'March Outstanding', 'April Outstanding', 'May Outstanding', 'June Outstanding', 'July Outstanding', 'August Outstanding', 'September Outstanding', 'October Outstanding', 'November Outstanding', 'December Outstanding']


            # Show plot
            st.plotly_chart(fig2)

            # Print DataFrame
            st.dataframe(df_Lower, height=150, use_container_width=True, hide_index=True)


            # Convert the chart to an image
            image = fig2.to_image(format="png", width=1200, height=550, scale=2.0)

            # Save the image to a file
            image_path = r"Resources\Plots\lower_put_options.png"
            with open(image_path, "wb") as f:
                f.write(image)
            # Set up the file name
            filename = "plotly_chart.png"
            # Convert the image to bytes
            image_bytes = io.BytesIO(image)

            

            # Trigger the download
            st.download_button(label="Download Image", data=image_bytes, file_name=filename, mime="image/png", key="download_button_2")

    else:
        st.write('no data')

    st.divider()

     # Display PCHP Data
    st.title("MTM Evaluation and Excel")

    # Create a formatted copy of the filtered DataFrame to preserve the original data
    formatted_df = filtered_df.copy()
    formatted_df_option = df_selected_sheet.copy()


        # Function to get list of months between two dates
    def get_months_between_dates(start_date, end_date):
        months = pd.date_range(start=start_date, end=end_date, freq='MS').strftime('%B').tolist()
        return months

    # Format date columns for better readability
    date_columns = ['FO.TradeDate', 'FO.StartFixDate', 'FO.EndFixDate', 'FO.Settlement_DeliveryDate']
    for column in date_columns:
        formatted_df[column] = formatted_df[column].dt.strftime('%d %b %Y')

    # Create a new column 'MonthsBetween' containing list of months between start and end fix dates
    formatted_df['Tenure'] = formatted_df.apply(lambda row: get_months_between_dates(row['FO.StartFixDate'], row['FO.EndFixDate']), axis=1)

    # Assuming 'formatted_df' is your DataFrame
    formatted_df.rename(columns={'Row Labels':'Trade Number'}, inplace=True)

    # Define a function to determine the option structure
    def get_option_structure(row):
        if row['FO.StrikePrice1'] != 0 and row['FO.StrikePrice2'] == 0:
            return 'Vanilla Options'
        elif row['FO.StrikePrice1'] != 0 and row['FO.StrikePrice2'] != 0:
            return 'Put Spreads Options'
        else:
            return 'Unknown'

    # Add a new column 'OptionStructure'
    formatted_df['OptionStructure'] = formatted_df.apply(lambda row: get_option_structure(row), axis=1)


    # Add a new column 'OptionStructure'
    formatted_df['OptionStructure'] = formatted_df.apply(lambda row: get_option_structure(row), axis=1)

    # Specify columns to display in the table
    columns_to_display = ['Trade Number','Portfolio','FO.TradeDate','FO.DealerID', 'FO.CounterpartyName','FO.OptionTypeLabel','OptionStructure','FO.PremiumStrike1', 'FO.PremiumStrike2','FO.NetPremium', 'FO.Position_Quantity',
                        'FO.StrikePrice1', 'FO.StrikePrice2', 'FO.StartFixDate', 'FO.EndFixDate','Tenor', 'FO.Settlement_DeliveryDate',
                        'O.January','O.February','O.March','O.April','O.May','O.June','O.July',
                        'O.August','O.September', 'O.October','O.November','O.December']

    # Reset index to start from 1
    formatted_df = formatted_df.reset_index(drop=True)

    # Start index from 1
    formatted_df.index = formatted_df.index + 1
    formatted_df_option_E = formatted_df.copy()

    # Find common months between both DataFrames
    common_months = [col for col in formatted_df_option.columns if col.startswith('O.')]

    # Iterate over each row in formatted_df
    for index, row in formatted_df.iterrows():
        # Get the Strike Price from formatted_df
        strike_price_1 = row['FO.StrikePrice1']
        strike_price_2 = row['FO.StrikePrice2']

        # Find corresponding row in formatted_df_option
        option_row_1 = formatted_df_option[formatted_df_option['Strike Price'] == strike_price_1]
        option_row_2 = formatted_df_option[formatted_df_option['Strike Price'] == strike_price_2]

        # Check if option_row is not empty
        if not option_row_1.empty and not option_row_2.empty:
            # Multiply the values in common months and update the row in formatted_df
            for month in common_months:
                formatted_df.at[index, month] = formatted_df.at[index, month] * round(option_row_1[month].iloc[0],3) - formatted_df.at[index, month] * round(option_row_2[month].iloc[0],3) 
               

    # Now the values in formatted_df are updated according to the conditions specified

    # List of columns related to the months
    month_columns = ['O.January','O.February','O.March','O.April','O.May','O.June','O.July',
                        'O.August','O.September', 'O.October','O.November','O.December']
    month_columns_value = ['January,USD', 'February,USD', 'March,USD', 'April,USD', 'May,USD', 'June,USD', 'July,USD', 'August,USD', 'September,USD', 'October,USD', 'November,USD', 'December,USD']
    month_columns_bbls = ['January,bbls', 'February,bbls', 'March,bbls', 'April,bbls', 'May,bbls', 'June,bbls', 'July,bbls', 'August,bbls', 'September,bbls', 'October,bbls', 'November,bbls', 'December,bbls']

    # Assuming 'formatted_df' is your DataFrame
    formatted_df.rename(columns={
        'O.January': 'January,USD',
        'O.February': 'February,USD',
        'O.March': 'March,USD',
        'O.April': 'April,USD',
        'O.May': 'May,USD',
        'O.June': 'June,USD',
        'O.July': 'July,USD',
        'O.August': 'August,USD',
        'O.September': 'September,USD',
        'O.October': 'October,USD',
        'O.November': 'November,USD',
        'O.December': 'December,USD'
    }, inplace=True)


    
    # Create new columns with default value 0 in formatted_df
    for col in month_columns:
        formatted_df[col] = 0

    # Assign values from formatted_df_option_E to formatted_df
    formatted_df[month_columns] = formatted_df_option_E[month_columns]  
    columns_to_display.extend(month_columns_value)

    formatted_df.rename(columns={
        'O.January': 'January,bbls',
        'O.February': 'February,bbls',
        'O.March': 'March,bbls',
        'O.April': 'April,bbls',
        'O.May': 'May,bbls',
        'O.June': 'June,bbls',
        'O.July': 'July,bbls',
        'O.August': 'August,bbls',
        'O.September': 'September,bbls',
        'O.October': 'October,bbls',
        'O.November': 'November,bbls',
        'O.December': 'December,bbls'
    }, inplace=True)

    columns_to_display = [f"{column.split('.')[1]},bbls" if column.startswith('O.') else column for column in columns_to_display]


    # Assuming you have a DataFrame named 'data' containing your dataset
    formatted_df['Value at inception, USD'] = formatted_df['FO.NetPremium'] * formatted_df['FO.Position_Quantity']
    columns_to_display.append('Value at inception, USD')

    # Create an empty list to store header names containing non-zero values
    header_names = []

    # Mapping dictionary for renaming column headers
    column_mapping = {
        'January,USD': 'O.January',
        'February,USD': 'O.February',
        'March,USD': 'O.March',
        'April,USD': 'O.April',
        'May,USD': 'O.May',
        'June,USD': 'O.June',
        'July,USD': 'O.July',
        'August,USD': 'O.August',
        'September,USD': 'O.September',
        'October,USD': 'O.October',
        'November,USD': 'O.November',
        'December,USD': 'O.December'
    }

    def col_round(x, decimals=3):
        if np.isnan(x):
            return x  # Return NaN as is
        factor = 10 ** decimals
        x *= factor
        frac = x - math.floor(x)
        if frac < 0.5:
            result = math.floor(x)
        else:
            result = math.ceil(x)
        return result / factor

    # Iterate through each row in the DataFrame
    for index, row in formatted_df.iterrows():
        # Initialize a list to store values for the current row
        selected_values = []
        
        # Extract FO.StrikePrice1 from formatted_df
        strike_price = row['FO.StrikePrice1']
        
        # Find the corresponding row(s) in df_selected_sheet based on Strike Price
        selected_rows = df_selected_sheet[df_selected_sheet['Strike Price'] == strike_price]

        selected_row_value = []
        # If corresponding rows are found
        if not selected_rows.empty:
            # Iterate through each row in selected_rows
            for selected_index, selected_row in selected_rows.iterrows():
                # Iterate through each column in month_columns_value
                for col in month_columns_value:
                    # Check if there's any value in the current column
                    if not pd.isnull(row[col]):
                        # If there's a value, add the corresponding column name from column_mapping to the list
                        selected_values.append(column_mapping[col])

                try:
                    # Extract values from selected_row for columns present in selected_values
                    selected_row_values = selected_row[selected_values]
                except KeyError:
                    # Handle the KeyError here
                    selected_row_values = None  # Or any other default value you want to return

                
               # Check if selected_row_values is not empty
                if selected_row_values is not None and len(selected_row_values) > 0:
                    # Calculate the average if multiple values exist
                    avg_selected_row_values = np.mean(selected_row_values)
                else:
                    # Handle the case when selected_row_values is empty
                    avg_selected_row_values = np.nan  # or any other default value or handling you prefer
                
                # Append the average to selected_row_value
                selected_row_value.append(avg_selected_row_values)

        

        # Assign the calculated average values to the 'Market Upper Premium' column in the DataFrame
        formatted_df.at[index, 'Market Upper Premium, USD'] = col_round(np.mean(selected_row_value),3)

    # Append the name 'Market Upper Premium' to columns_to_display
    columns_to_display.append('Market Upper Premium, USD')

    # Iterate through each row in the DataFrame
    for index, row in formatted_df.iterrows():
        # Initialize a list to store values for the current row
        selected_values = []
        
        # Extract FO.StrikePrice1 from formatted_df
        strike_price = row['FO.StrikePrice2']
        
        # Find the corresponding row(s) in df_selected_sheet based on Strike Price
        selected_rows = df_selected_sheet[df_selected_sheet['Strike Price'] == strike_price]

        selected_row_value = []
        # If corresponding rows are found
        if not selected_rows.empty:
            # Iterate through each row in selected_rows
            for selected_index, selected_row in selected_rows.iterrows():
                # Iterate through each column in month_columns_value
                for col in month_columns_value:
                    # Check if there's any value in the current column
                    if not pd.isnull(row[col]):
                        # If there's a value, add the corresponding column name from column_mapping to the list
                        selected_values.append(column_mapping[col])

                try:
                    # Extract values from selected_row for columns present in selected_values
                    selected_row_values = selected_row[selected_values]
                except KeyError:
                    # Handle the KeyError here
                    selected_row_values = None  # Or any other default value you want to return
                
                               # Check if selected_row_values is not empty
                if selected_row_values is not None and len(selected_row_values) > 0:
                    # Calculate the average if multiple values exist
                    avg_selected_row_values = round(np.mean(selected_row_values),3)
                else:
                    # Handle the case when selected_row_values is empty
                    avg_selected_row_values = np.nan  # or any other default value or handling you prefer
                
                # Append the average to selected_row_value
                selected_row_value.append(avg_selected_row_values)

        # Assign the calculated average values to the 'Market Upper Premium' column in the DataFrame
        formatted_df.at[index, 'Market Lower Premium, USD'] = col_round(np.mean(selected_row_value),3)

    # Append the name 'Market Upper Premium' to columns_to_display
    columns_to_display.append('Market Lower Premium, USD')

    # Assuming you have a DataFrame named 'data' containing your dataset
    formatted_df['Market Net Premium, USD'] = formatted_df['Market Upper Premium, USD'] - formatted_df['Market Lower Premium, USD']
    columns_to_display.append('Market Net Premium, USD')




    # Add a new column 'Total' containing the sum of values in the month columns
    formatted_df['Current Value, USD'] = formatted_df[month_columns_value].sum(axis=1)
    columns_to_display.append('Current Value, USD')



    # Extract unique items from 'FO.EndFixDate'
    unique_items = formatted_df['FO.EndFixDate'].unique()

    # Define the color sequence
    color_discrete_sequence_2 = [
        "#763f98",  # Original color - R118 G063 B152
        "#20419a",  # Original color - R032 G065 B154
        "#fdb924",
        "#00b1a9",  # Original color - R253 G185 B036
        "#bfd730",  # Original color - R191 G215 B048
        "#007b73",  # Shade of R000 G177 B169
        "#3a1d4c",  # Shade of R118 G063 B152
        "#101e4a",  # Shade of R032 G065 B154
        "#cc8b1c",  # Shade of R253 G185 B036
        "#8e9c1b",  # Shade of R191 G215 B048
        "#b62e20",  # Background color - Similar theme color
        "#ff6f61",  # Additional color 1
        "#4ecdc4",  # Additional color 2
        "#ff9f51",  # Additional color 3
        "#2ab7ca"   # Additional color 4
    ]


    # Create a dictionary with unique items as keys and corresponding colors
    Month_colors = dict(zip(unique_items, color_discrete_sequence_2[:len(unique_items)]))

    st.subheader("Mark to Market result")
    total_sum = formatted_df['Current Value, USD'].sum()
    total_sum_incep = formatted_df['Value at inception, USD'].sum()

    # Assuming formatted_df is your DataFrame
    month_columns = [
        'January,bbls', 'February,bbls', 'March,bbls', 'April,bbls', 
        'May,bbls', 'June,bbls', 'July,bbls', 'August,bbls', 
        'September,bbls', 'October,bbls', 'November,bbls', 'December,bbls'
    ]

    # Check if there's a value in any of the month columns
    condition = formatted_df[month_columns].notnull().any(axis=1)

    # Sum the 'Value at inception, USD' where the condition is True
    total_outstanding_incep = formatted_df.loc[condition, 'Value at inception, USD'].sum()

    col5,col6,col7,col8 = st.columns(4)
    with col5:
        st.metric(label='Total Inception Value, USD', value=str(f" {total_sum_incep:,.0f} "))
    with col6:
        st.metric(label='Outstanding Inception Value, USD', value=str(f" {total_outstanding_incep:,.0f} "))
    with col7:
        st.metric(label='Current Oustanding Value, USD', value=str(f" {total_sum:,.0f} "))
    with col8:
        st.metric(label='MTM Movements, USD', value=str(f" {total_sum - total_outstanding_incep   :,.0f} "))
    
    
    # Calculate Volume executed versus Counterparty
    st.subheader("Current Option Value per Counterparty")

    # Create a checkbox for toggling colorization
    use_color = st.checkbox('Separation by Expiration Date', value=True)
    
    # Add a column for custom colors based on DealerID
    formatted_df['Color_2'] = formatted_df['FO.EndFixDate'].map(Month_colors)

    if use_color:
        # If the checkbox is checked, apply colorization
        fig1 = px.histogram(formatted_df, x='FO.Acronym', y='Current Value, USD', color='FO.EndFixDate', 
                           title='Value of Outstanding Volumes by expiration date', 
                           color_discrete_map=Month_colors)
        
         # Update the x-axis category order
        fig1.update_xaxes(categoryorder='total descending')

        # Rename x and y labels
        fig1.update_xaxes(title_text='Counterparties')
        fig1.update_yaxes(title_text='Value, USD')

        # Add values at the top of each bar
        fig1.update_traces(texttemplate='%{y:.2s}', textposition='inside')

        # Convert the chart to an image
        image = fig1.to_image(format="png", width=1200, height=600, scale=2.0)

        # Save the image to a file
        image_path = r"Resources\Plots\volume_active_1.png"
        with open(image_path, "wb") as f:
            f.write(image)
    else:
        # If the checkbox is unchecked, don't apply colorization
        fig1 = px.histogram(formatted_df, x='FO.Acronym', y='Current Value, USD', 
                           title='Total Value of Outstanding Volume',color_discrete_sequence=["#00b1a9"])
        
         # Update the x-axis category order
        fig1.update_xaxes(categoryorder='total descending')

        # Rename x and y labels
        fig1.update_xaxes(title_text='Counterparties')
        fig1.update_yaxes(title_text='Value, USD')

        # Add values at the top of each bar
        fig1.update_traces(texttemplate='%{y:.2s}', textposition='inside')

        # Convert the chart to an image
        image = fig1.to_image(format="png", width=1200, height=600, scale=2.0)

        # Save the image to a file
        image_path = r"Resources\Plots\volume_active_2.png"
        with open(image_path, "wb") as f:
            f.write(image)


    # Show the Plotly figure in Streamlit
    st.plotly_chart(fig1, use_container_width=True, height=200)

    

    # Now the values in formatted_df_option are updated according to the conditions specified
    with st.container():
        # Show the formatted DataFrame using st.dataframe
        st.dataframe(formatted_df[columns_to_display], height=500 ,use_container_width=True)

    # buffer to use for excel writer
    buffer = io.BytesIO()

    from openpyxl import Workbook
    from openpyxl.styles import NamedStyle, Alignment, PatternFill, Font

    # Download Button
    @st.cache_data
    def convert_to_excel(formatted_df, df_selected_sheet):
        # Create a buffer to hold the Excel file
        buffer = BytesIO()

        # Create Excel writer object
        with pd.ExcelWriter(buffer, engine='xlsxwriter', date_format='dd/mm/yyyy') as writer:

            # List of columns to apply date formatting
            date_columns = ['FO.TradeDate', 'FO.StartFixDate', 'FO.EndFixDate', 'FO.Settlement_DeliveryDate']

            # Loop through each column and apply formatting
            for col in date_columns:
            # Convert the text format to datetime format
                formatted_df[col] = pd.to_datetime(formatted_df[col], format='%d %b %Y', errors='coerce')

            # Write formatted_df to the first sheet
            formatted_df.to_excel(writer, sheet_name='Portfolio Sum', index=False)

            # Write df_selected_sheet to the second sheet
            df_selected_sheet.to_excel(writer, sheet_name='Option Data', index=False)

            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet1 = writer.sheets['Portfolio Sum']
            worksheet2 = writer.sheets['Option Data']

            # Define cell formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'border': 1,
                'font_color': 'white',  # Set font color to white
                'bg_color': '#38B09D'  # Set background color to 38B09D
            })
            data_format = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'border': 1})
            date_format = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'border': 1,'num_format': 'dd/mm/yyyy'})

            # List of columns to apply date formatting
            date_columns = ['FO.TradeDate', 'FO.StartFixDate', 'FO.EndFixDate', 'FO.Settlement_DeliveryDate']

            # Loop through each column and apply formatting
            for col in date_columns:
                if col in formatted_df.columns:
                    # Get column index
                    col_index = formatted_df.columns.get_loc(col)
                    
                    # Write formatted dates to the worksheet
                    for row_num in range(1, len(formatted_df) + 1):
                        date_value = formatted_df.iloc[row_num - 1][col]
                        worksheet1.write(row_num, col_index, date_value, date_format)
                

            # Apply formatting to first sheet (Portfolio Sum)
            for col_num, value in enumerate(formatted_df.columns.values):
                worksheet1.write(0, col_num, value, header_format)
                worksheet1.set_column(col_num, col_num, 15, data_format)  # Set column width to 15

            # Apply formatting to second sheet (Option Data)
            for col_num, value in enumerate(df_selected_sheet.columns.values):
                worksheet2.write(0, col_num, value, header_format)
                worksheet2.set_column(col_num, col_num, 15, data_format)  # Set column width to 15

            # Set row height in points (1 point ≈ 0.75 pixels)
            row_height_in_points = 50
            worksheet1.set_default_row(row_height_in_points)  # Set default row height for the first sheet
            worksheet2.set_default_row(row_height_in_points)  # Set default row height for the second sheet

            # Freeze the header row
            worksheet1.freeze_panes(1, 0)  # Freeze the first row in the first sheet

        return buffer.getvalue()


    excel_data = convert_to_excel(formatted_df[columns_to_display], df_selected_sheet)

    # Get today's date and format it
    today_date = datetime.now().strftime('%d-%m-%Y')

    # Download button to download Excel file
    download_button = st.download_button(
        label="Download data as Excel",
        data=excel_data,
        file_name=f"{today_date}_MTM.xlsx",  # Set file name with today's date
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

with tabITM:
    st.title("In-The-Money Data")
    st.subheader("Average Monthly Brent Price")

    # List of months
    months = ["January", "February", "March", "April", "May", "June", 
            "July", "August", "September", "October", "November", "December"]

    # Add "Actual" columns
    for month in months:
        ITM_df[f"Actual {month}"] = None
    

    # Cache the loading of the Excel file
    @st.cache_data
    def load_data(file, sheet):
        return pd.read_excel(file, sheet).round(3)

    # Load the default dataset
    default_file = "PCHP Data.xlsx"
    default_sheet = "Actual_Average_Brent"
    Actual_Average_Brent_df = load_data(default_file, default_sheet)

    ## Let the user know they can edit the default dataset
    st.write("You can edit the data in the table below. The default data is preloaded.")

    # Editable table with default data preloaded
    Actual_Average_Brent_df = st.data_editor(
        Actual_Average_Brent_df,
        num_rows="dynamic",
        column_config={
            "Year": st.column_config.TextColumn("Year"),
            **{col: st.column_config.NumberColumn(col, format="%.3f") for col in Actual_Average_Brent_df.columns if col != "Year"}
        },use_container_width=True
    )

    # Ensure the DataFrame is rounded to two decimal places
    Actual_Average_Brent_df = Actual_Average_Brent_df.round(3)

    # Display the final DataFrame
    #st.write("### Final Data:")
    #st.dataframe(Actual_Average_Brent_df, use_container_width=True)

    # Create a mapping dictionary from Actual_Average_Brent_df
    actuals_dict = {}
    for index, row in Actual_Average_Brent_df.iterrows():
        actuals_dict[row['Year']] = row.drop('Year').to_dict()

    # Now, populate the 'Actual' columns in ITM_df
    months = ["January", "February", "March", "April", "May", "June", 
            "July", "August", "September", "October", "November", "December"]

    for month in months:
        column_name = f"Actual {month}"
        ITM_df[column_name] = ITM_df['Portfolio'].apply(
            lambda x: actuals_dict.get(x, {}).get(month, None)
        )

    # Updated list of months with _ITM
    months_ITM = ["January_ITM", "February_ITM", "March_ITM", "April_ITM", "May_ITM", "June_ITM",
                "July_ITM", "August_ITM", "September_ITM", "October_ITM", "November_ITM", "December_ITM"]

    # Fill the DataFrame with calculated values for each month
    for month in months_ITM:
        actual_month = 'Actual ' + month.split('_')[0]
        month_value = month.split('_')[0]
        
        ITM_df[month] = ((ITM_df[actual_month] < ITM_df['FO.StrikePrice1']).astype(int) * (ITM_df['FO.StrikePrice1'] - ITM_df[actual_month]) * ITM_df[month_value])

    # List of columns to remove
    columns_to_remove = [
        'Unnamed: 32', 'O.January', 'O.February', 'O.March', 'O.April', 'O.May', 
        'O.June', 'O.July', 'O.August', 'O.September', 'O.October', 'O.November', 
        'O.December', 'Total Outstanding'
    ]

    # Remove the columns from ITM_df
    ITM_df = ITM_df.drop(columns=columns_to_remove, errors='ignore')

    # Select only the FO.Acronym and *_ITM columns
    columns_to_plot = ['FO.Acronym'] + [col for col in ITM_df.columns if col.endswith('_ITM')]

    # Filter the DataFrame to include only the relevant columns
    filtered_df = ITM_df[columns_to_plot]

   # Melt the DataFrame to long format
    ITM_long = filtered_df.melt(id_vars='FO.Acronym', 
                                var_name='Month', 
                                value_name='Value')

    # Clean up the month names (removing '_ITM')
    ITM_long['Month'] = ITM_long['Month'].str.replace('_ITM', '')

    # Group by Month and FO.Acronym and sum the values
    ITM_grouped = ITM_long.groupby(['Month', 'FO.Acronym'], as_index=False)['Value'].sum()

    # Define the custom month order
    month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                'July', 'August', 'September', 'October', 'November', 'December']

    # Ensure 'Month' is a categorical type with the defined order
    ITM_grouped['Month'] = pd.Categorical(ITM_grouped['Month'], categories=month_order, ordered=True)

    # Create the stacked bar chart with the custom order
    fig = px.bar(ITM_grouped, 
                x='Month', 
                y='Value', 
                color='FO.Acronym', 
                title='Monthly ITM Distribution by Counterparty',
                labels={'Month': 'Month', 'Value': 'ITM Value', 'FO.Acronym': 'Acronym'},
                barmode='stack',
                category_orders={'Month': month_order})  # Enforce custom month order

    # Add values at the top of each bar
    fig.update_traces(texttemplate='%{y}', textposition='inside')

    # Display in Streamlit
    st.plotly_chart(fig, use_container_width=True)

    # Pivot the DataFrame to reshape it
    pivot_df = ITM_grouped.pivot(index='FO.Acronym', columns='Month', values='Value')

    # Fill NaN values with 0 (optional)
    pivot_df = pivot_df.fillna(0)

    # Add a "Total" column for each row (horizontal sum)
    pivot_df['Total'] = pivot_df.sum(axis=1)

    # Add a "Total" row for each column (vertical sum)
    pivot_df.loc['Total'] = pivot_df.sum()

    # Display the pivot table in Streamlit
    st.dataframe(pivot_df, use_container_width=True,height=500)




with tab4:
    
    
    def create_download_link(val, filename):
        b64 = base64.b64encode(val)  # val looks like b'...'
        return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{filename}.pdf">Download file</a>'

    def create_letterhead(pdf, WIDTH):
        pdf.image(r"Resources/Blue Modern Business Letterhead.jpg", 0, 0, WIDTH)

    def create_title(title, pdf):
        # Add main title
        pdf.set_font('Helvetica', 'b', 20)  
        pdf.ln(40)
        pdf.write(5, title)
        pdf.ln(10)
        # Add date of report
        pdf.set_font('Helvetica', '', 14)
        pdf.set_text_color(r=128,g=128,b=128)
        today = time.strftime("%d/%m/%Y")
        pdf.write(4, f'{today}')
        # Add line break
        pdf.ln(10)

    def write_to_pdf(pdf, words):
        # Set text colour, font size, and font type
        pdf.set_text_color(r=0,g=0,b=0)
        pdf.set_font('Helvetica', '', 12)
        pdf.write(5, words)

    class PDF(FPDF):
        def footer(self):
            self.set_y(-15)
            self.set_font('Helvetica', 'I', 8)
            self.set_text_color(128)
            self.cell(0, 10, 'Page ' + str(self.page_no()), 0, 0, 'C')

    # Global Variables
    TITLE = "Portfolio Commodity Hedging Program Report"
    WIDTH = 210
    HEIGHT = 297

    # Create PDF
    pdf = PDF() # A4 (210 by 297 mm)

    # Add Page
    pdf.add_page()

    # Add lettterhead and title
    create_letterhead(pdf, WIDTH)
    create_title(TITLE, pdf)
    pdf.image(r"Resources\Plots\Brent.png", x=5, y=pdf.get_y(), w=200)
    
    pdf.image(r"Resources\Plots\volume_executed.png", x=5, y=150, w=200)
    pdf.image(r"Resources\Plots\Execution_table.png", x=5, y=200, w=200)
    
    pdf.ln(100)
    
    #add page_2
    pdf.add_page()
    create_letterhead(pdf,WIDTH)
    create_title(TITLE,pdf)
    pdf.image(r"Resources\Plots\volume_cp.png", x=5, y=pdf.get_y(), w=200)
    pdf.image(r"Resources\Plots\volume_dealer.png", x=5, y=150, w=200)
    

    # Add Page
    pdf.add_page()

    # Add lettterhead
    create_letterhead(pdf, WIDTH)
    create_title("MARK TO MARKET REPORT", pdf)

    # Add dynamically generated image to the PDF
    # Assuming image_bytes contains the bytes of the image generated by Plotly
    
    pdf.image(r"Resources\Plots\Outstanding_FO.StrikePrice1.png", x=5, y=pdf.get_y(), w=100)
    pdf.image(r"Resources\Plots\Outstanding_FO.StrikePrice2.png", x=105, y=pdf.get_y(), w=100)
    pdf.image(r"Resources\Plots\upper_put_options.png", x=5, y=130, w=100)

    pdf.image(r"Resources\Plots\lower_put_options.png", x=105, y=130, w=100)

    pdf.image(r"Resources\Plots\volume_active_1.png", x=5, y=180, w=100)
    

    # Generate the PDF and provide download link
    pdf_output = pdf.output(dest="S").encode("latin-1")
    html = create_download_link(pdf_output, "report")
    st.markdown(html, unsafe_allow_html=True)

