import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import re
from io import BytesIO

st.set_page_config(page_title="Invoice Extractor", layout="centered")

st.title("📄 Atomize Invoice Extractor")
st.write("Upload the `.html` file.")

# Upload do HTML
html_file = st.file_uploader("📤 Faça upload do arquivo HTML", type=["html"])

def extract_all_invoices_excluding_below_threshold(html_content, invoice_threshold=14728):
    soup = BeautifulSoup(html_content, 'html.parser')

    columns = [
        'Period Allocation Account',
        'Created from',
        'Voucher text',
        'Original amount',
        'Account',
        'Period allocation amount',
        'Period',
        'Sum completed',
        'Sum not completed'
    ]
    all_data = []

    data_rows = soup.find_all('tr', attrs={'valign': 'top', 'style': 'word-wrap:break-word;'})

    for row in data_rows:
        cells = row.find_all('td')
        if len(cells) < 9:
            continue

        created_from = extract_created_from(cells[0])

        invoice_number_match = re.search(r'(\d+)', created_from)
        if invoice_number_match and int(invoice_number_match.group(1)) < invoice_threshold:
            continue

        parent = row
        account_number = None
        for _ in range(10):
            parent = parent.find_parent()
            if parent:
                b_tag = parent.find('b', string=re.compile(r'Period allocation account:\s*\d+'))
                if b_tag:
                    match = re.search(r'Period allocation account:\s*(\d+)', b_tag.get_text())
                    if match:
                        account_number = int(match.group(1))
                        break

        voucher_text = extract_voucher_text_from_cell(cells[1])
        original_amount = clean_amount(cells[2].get_text(strip=True))
        account = clean_text(cells[3].get_text(strip=True))
        period_alloc = clean_amount(cells[4].get_text(strip=True))
        period = clean_text(cells[5].get_text(strip=True))
        sum_completed = clean_amount(cells[7].get_text(strip=True))
        sum_not_completed = clean_amount(cells[8].get_text(strip=True))

        all_data.append([
            account_number,
            created_from,
            voucher_text,
            original_amount,
            account,
            period_alloc,
            period,
            sum_completed,
            sum_not_completed
        ])

    return pd.DataFrame(all_data, columns=columns)

# Helpers
def extract_created_from(cell):
    link = cell.find('a')
    if link:
        created_from = link.get_text().strip()
    else:
        created_from = cell.get_text().strip()
    return created_from.split('\n')[0].strip()

def extract_voucher_text_from_cell(cell):
    cell_html = str(cell).replace('<br>', ' | ').replace('<br/>', ' | ')
    temp_soup = BeautifulSoup(cell_html, 'html.parser')
    return temp_soup.get_text().strip()

def clean_amount(text):
    return text.replace('\xa0', '').replace(' ', '').strip()

def clean_text(text):
    return text.replace('\xa0', '').strip()

if html_file:
    content = html_file.read().decode("utf-8")
    df = extract_all_invoices_excluding_below_threshold(content)

    if not df.empty:
        # Criar df_invoice com base no df extraído
        df_invoice = pd.DataFrame()
        df_invoice['Invoice No.'] = df['Created from']
        df_invoice['Parent/Customer No.'] = df['Account']
        df_invoice['Description'] = df['Voucher text']
        df_invoice['Unit Price Excl. VAT'] = df['Period allocation amount']
        df_invoice['CUSTOMER Dimension'] = df['Account']
        df_invoice['No.'] = "RMS PACKAGE"
        df_invoice['Quantity'] = "1"
        df_invoice['Deferral Start Date'] = df['Period'].apply(lambda x: x.split(' - ')[0])
        df_invoice['Deferral End Date'] = df['Period'].apply(lambda x: x.split(' - ')[1])

        new_columns = [
            "Subaccount No.", "Document Date", "Posting Date", "Due Date",
            "VAT Date", "Currency Code", "Type",
            "VAT Prod. Posting Group", "Deferral Code", "BU Dimension", "C Dimension",
            "ENTITY Dimension", "IC Dimension", "PRICE Dimension", "PRODUCT Dimension",
            "RECURRENCE Dimension", "SUBPRODUCT Dimension", "TAX DEDUCTIBILITY Dimension",
            "Reseller Code", "Apply Overpayments"
        ]
        for col in new_columns:
            df_invoice[col] = ""

        ordered_columns = [
            "Invoice No.", "Parent/Customer No.", "Subaccount No.", "Document Date", "Posting Date", "Due Date",
            "VAT Date", "Currency Code", "Type", "No.", "Description", "Quantity", "Unit Price Excl. VAT",
            "VAT Prod. Posting Group", "Deferral Code", "Deferral Start Date", "Deferral End Date",
            "BU Dimension", "C Dimension", "ENTITY Dimension", "IC Dimension", "PRICE Dimension",
            "PRODUCT Dimension", "RECURRENCE Dimension", "SUBPRODUCT Dimension",
            "TAX DEDUCTIBILITY Dimension", "CUSTOMER Dimension", "Reseller Code", "Apply Overpayments"
        ]
        df_invoice = df_invoice[ordered_columns]

        # Gerar arquivo Excel para download
        output = BytesIO()
        df_invoice.to_excel(output, index=False, engine='openpyxl')
        st.success("✅ Sucess!")

        st.download_button(
            label="📥 Download Excel",
            data=output.getvalue(),
            file_name="Atomize_invoice.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No valid invoice invoice found. All under 14728")
