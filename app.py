import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import io

# Function to convert Excel serial date to datetime
def excel_serial_to_date(serial):
    if pd.isna(serial) or not isinstance(serial, (int, float)):
        return None
    base_date = datetime(1899, 12, 30)
    return base_date + timedelta(days=serial)

# Function to generate report for a single date
def generate_total_geral_report(row, target_date_str):
    beach = row[1] if not pd.isna(row[1]) else 0
    norte = row[2] if not pd.isna(row[2]) else 0
    torre = row[3] if not pd.isna(row[3]) else 0
    adicionais = row[4] if not pd.isna(row[4]) else 0
    total = row[5] if not pd.isna(row[5]) else 0
    
    report = f"Relatório para a data {target_date_str}:\n"
    report += f"Beach: {beach}\n"
    report += f"Norte: {norte}\n"
    report += f"Torre: {torre}\n"
    report += f"Adicionais: {adicionais}\n"
    report += f"Total: {total}\n"
    return report, beach, norte, torre, adicionais, total

# Function to generate PDF for Beach Plaza (fit to one page)
def generate_beach_plaza_pdf(df_beach):
    fig, ax = plt.subplots(figsize=(12, 8))  # Adjusted for better fit
    ax.axis('tight')
    ax.axis('off')
    table = ax.table(cellText=df_beach.values, colLabels=df_beach.columns, loc='center', cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(5)  # Smaller font for large sheets
    table.scale(1.5, 1.5)  # Scale up for readability
    pdf_buffer = io.BytesIO()
    with PdfPages(pdf_buffer) as pdf:
        pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    pdf_buffer.seek(0)
    return pdf_buffer

# Streamlit config
st.set_page_config(page_title="Dashboard Hospedagem MRV", layout="wide")
st.title("Dashboard Interativo - Hospedagem MRV")

# Load Excel (with error handling)
file_path = 'HOSPEDAGEM MRV271125.xlsx'
try:
    df_total_geral = pd.read_excel(file_path, sheet_name='TOTAL GERAL', header=None)
    df_beach_plaza = pd.read_excel(file_path, sheet_name='Beach Plaza')
    df_beach_plaza = df_beach_plaza.fillna('')
except FileNotFoundError:
    st.error("Arquivo Excel não encontrado. Certifique-se de que 'HOSPEDAGEM MRV271125.xlsx' está no diretório.")
    st.stop()

# Extract available dates from TOTAL GERAL (rows 3-11, column A)
df_total_geral = df_total_geral.iloc[3:12].reset_index(drop=True)
available_dates = []
for _, row in df_total_geral.iterrows():
    serial = row[0]
    date_obj = excel_serial_to_date(serial)
    if date_obj:
        available_dates.append(date_obj.strftime('%Y-%m-%d'))

# Sidebar for multi-date selection
st.sidebar.header("Filtros")
selected_dates = st.sidebar.multiselect("Selecione as datas desejadas:", sorted(set(available_dates)))

# Generate and display reports
if selected_dates:
    st.subheader("Relatórios TOTAL GERAL")
    reports = []
    sum_beach, sum_norte, sum_torre, sum_adicionais, sum_total = 0, 0, 0, 0, 0
    
    for target_date_str in sorted(selected_dates):
        target_date = datetime.strptime(target_date_str, '%Y-%m-%d')
        for _, row in df_total_geral.iterrows():
            date_obj = excel_serial_to_date(row[0])
            if date_obj and date_obj.date() == target_date.date():
                report, beach, norte, torre, adicionais, total = generate_total_geral_report(row, target_date_str)
                reports.append(report)
                sum_beach += beach
                sum_norte += norte
                sum_torre += torre
                sum_adicionais += adicionais
                sum_total += total
                break
        else:
            st.warning(f"Data {target_date_str} não encontrada no Excel.")
    
    # Display individual reports
    for report in reports:
        st.text(report)
    
    # Display sum if multiple dates
    if len(selected_dates) > 1:
        total_report = f"Soma Total para as datas selecionadas:\n"
        total_report += f"Beach: {sum_beach}\n"
        total_report += f"Norte: {sum_norte}\n"
        total_report += f"Torre: {sum_torre}\n"
        total_report += f"Adicionais: {sum_adicionais}\n"
        total_report += f"Total: {sum_total}\n"
        st.subheader("Soma Total")
        st.text(total_report)
    
    # Download combined report as PDF
    combined_report = "\n\n".join(reports)
    if len(selected_dates) > 1:
        combined_report += "\n\n" + total_report
    fig_text, ax_text = plt.subplots(figsize=(8, 11))
    ax_text.text(0.1, 0.9, combined_report, fontsize=10, va='top')
    ax_text.axis('off')
    pdf_buffer_text = io.BytesIO()
    with PdfPages(pdf_buffer_text) as pdf:
        pdf.savefig(fig_text, bbox_inches='tight')
    plt.close()
    pdf_buffer_text.seek(0)
    st.download_button("Baixar Relatório Combinado como PDF", data=pdf_buffer_text, file_name="total_geral_report.pdf", mime="application/pdf")

# Display full Beach Plaza sheet
st.subheader("Aba Beach Plaza (Completa)")
st.dataframe(df_beach_plaza, use_container_width=True)

# Download Beach Plaza as PDF
pdf_buffer_beach = generate_beach_plaza_pdf(df_beach_plaza)
st.download_button("Baixar Beach Plaza como PDF", data=pdf_buffer_beach, file_name="beach_plaza_report.pdf", mime="application/pdf")

# Hide Streamlit defaults
hide_style = """
<style>
#MainMenu {visibility:hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_style, unsafe_allow_html=True)
