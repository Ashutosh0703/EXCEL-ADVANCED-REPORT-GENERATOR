import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title='Excel Report Generator', layout='centered')
st.title('📊  Excel Advance Report Generator')

st.markdown("""
**Instructions:**
1. Upload your sales CSV file.
2. Click 'Generate Excel Report'.
3. Download your full Excel file.
""")

uploaded_file = st.file_uploader("Upload your sales CSV", type="csv")

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.subheader("Data Preview")
    st.dataframe(df.head(), use_container_width=True)

    if st.button("Generate Excel Report"):
        # Data summaries
        df['Date'] = pd.to_datetime(df['Date'])
        region_summary = df.groupby('Region')['Net_Sales'].sum().reset_index().sort_values('Net_Sales', ascending=False)
        product_summary = df.groupby('Product')['Net_Sales'].sum().reset_index().sort_values('Net_Sales', ascending=False)
        monthly_summary = df.groupby(df['Date'].dt.to_period('M'))['Net_Sales'].sum().reset_index()
        monthly_summary['Date'] = monthly_summary['Date'].astype(str)

        # Write Excel to memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            title_fmt = workbook.add_format({'font_size': 16, 'bold': True, 'align': 'center', 'font_color': '#1F4E79', 'bg_color': '#D9E2F3'})
            hdr_fmt   = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#4472C4', 'align': 'center'})
            cur_fmt   = workbook.add_format({'num_format': ' #,##0', 'align': 'right'})
            num_fmt   = workbook.add_format({'num_format': '#,##0', 'align': 'center'})

            # Raw Data Sheet
            df.to_excel(writer, index=False, sheet_name='Raw Data', startrow=1)
            ws1 = writer.sheets['Raw Data']
            ws1.merge_range('A1:N1', 'Raw Sales Data', title_fmt)
            for i, col in enumerate(df.columns):
                ws1.write(1, i, col, hdr_fmt)
                ws1.set_column(i, i, 15)

            # Region Summary Sheet
            region_summary.to_excel(writer, index=False, sheet_name='Region Summary', startrow=2)
            ws2 = writer.sheets['Region Summary']
            ws2.merge_range('A1:C1', 'Total Sales by Region', title_fmt)
            for i, col in enumerate(region_summary.columns):
                ws2.write(2, i, col, hdr_fmt)
                ws2.set_column(i, i, 18)
            ws2.conditional_format(3, 1, 3 + len(region_summary) - 1, 1,
                                   {'type': '3_color_scale', 'min_color': '#F8696B', 'mid_color': '#FFEB84', 'max_color': '#63BE7B'})
            chart1 = workbook.add_chart({'type': 'column'})
            chart1.add_series({
                'name': 'Net Sales',
                'categories': ['Region Summary', 3, 0, 3+len(region_summary)-1, 0],
                'values':     ['Region Summary', 3, 1, 3+len(region_summary)-1, 1],
                'fill':       {'color': '#4472C4'}
            })
            chart1.set_title({'name': 'Net Sales by Region'})
            chart1.set_y_axis({'name': 'Net Sales ( )'})
            ws2.insert_chart('E3', chart1)

            # Product Summary Sheet
            product_summary.to_excel(writer, index=False, sheet_name='Product Summary', startrow=2)
            ws3 = writer.sheets['Product Summary']
            ws3.merge_range('A1:C1', 'Market Share by Product', title_fmt)
            for i, col in enumerate(product_summary.columns):
                ws3.write(2, i, col, hdr_fmt)
                ws3.set_column(i, i, 18)
            ws3.conditional_format(3, 1, 3 + len(product_summary) - 1, 1,
                                   {'type': 'data_bar', 'bar_color': '#70AD47', 'bar_solid': True})
            chart2 = workbook.add_chart({'type': 'pie'})
            chart2.add_series({
                'name': 'Market Share',
                'categories': ['Product Summary', 3, 0, 3+len(product_summary)-1, 0],
                'values':     ['Product Summary', 3, 1, 3+len(product_summary)-1, 1]
            })
            chart2.set_title({'name': 'Market Share by Product'})
            ws3.insert_chart('E3', chart2)

            # Monthly Trends Sheet
            monthly_summary.to_excel(writer, index=False, sheet_name='Monthly Trends', startrow=2)
            ws4 = writer.sheets['Monthly Trends']
            ws4.merge_range('A1:B1', 'Monthly Sales Trends', title_fmt)
            for i, col in enumerate(monthly_summary.columns):
                ws4.write(2, i, col, hdr_fmt)
                ws4.set_column(i, i, 18)
            chart3 = workbook.add_chart({'type': 'line'})
            chart3.add_series({
                'name': 'Net Sales',
                'categories': ['Monthly Trends', 3, 0, 3+len(monthly_summary)-1, 0],
                'values':     ['Monthly Trends', 3, 1, 3+len(monthly_summary)-1, 1],
                'line':       {'color': '#4472C4'}
            })
            chart3.set_title({'name': 'Sales Trend by Month'})
            chart3.set_y_axis({'name': 'Net Sales ( )'})
            ws4.insert_chart('D3', chart3)

            # Executive Summary Sheet
            summary_sheet = workbook.add_worksheet('Executive Summary')
            summary_sheet.merge_range('A1:D1', 'Sales Executive Summary', title_fmt)
            summary_sheet.write('A2', 'Generated:', hdr_fmt)
            summary_sheet.write('B2', datetime.now().strftime('%Y-%m-%d %H:%M'))
            summary_sheet.write('A4', 'Total Net Sales', hdr_fmt)
            summary_sheet.write('B4', df['Net_Sales'].sum(), cur_fmt)
            summary_sheet.write('A5', 'Total Units Sold', hdr_fmt)
            summary_sheet.write('B5', df['Units_Sold'].sum(), num_fmt)
            summary_sheet.write('A6', 'Top Region', hdr_fmt)
            summary_sheet.write('B6', region_summary.iloc[0]['Region'])
            summary_sheet.write('A7', 'Top Product', hdr_fmt)
            summary_sheet.write('B7', product_summary.iloc[0]['Product'])
            summary_sheet.set_column('A:B', 24)

        output.seek(0)
        st.success("✅ Excel report generated!")
        st.download_button(
            label="⬇️ Download Excel Report",
            data=output,
            file_name="sales_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload a CSV file to begin.")
