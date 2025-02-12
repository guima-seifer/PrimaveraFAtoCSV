import pdfplumber
import pandas as pd
import re
import glob

# Path to the PDF files
pdf_path_pattern = "pdfs/*.pdf"

# Function to extract invoice data
def extract_invoice_data(pdf_path):
    data = {
        "Invoice Number": "",
        "Date": "",
        "Currency": "",
        "Exchange Rate": "",
        "Payment Due": "",
        "Commercial Discount": "",
        "Payment Terms": "",
        "Customer": "",
        "Address": "",
        "Client_NIF": "",
        "Total Amount": "",
        "Total Liquid": "",
        "Execution Location": "",
        "VAT": "",
        "Items": []
    }
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
            # print(f"Extracted Text from {pdf_path}:\n", text)
    except Exception as e:
        print(f"Failed to read {pdf_path}: {e}")
        return data

    try:
        # Extract invoice number
        match = re.search(r"Fatura FT ([\w/.]+)", text)
        if match:
            data["Invoice Number"] = match.group(1)
        
        # Extract date
        match = re.search(r"(\d{4}-\d{2}-\d{2})", text)
        if match:
            data["Date"] = match.group(1)

        # Extract currency
        match = re.search(r"Data\s*(.+?)\n", text)
        if match:
            info = match.group(1).split(" ")
            # print(info)
            data["Client_NIF"] = info[0]
            data["Currency"] = info[1]
            # Extract exchange rate
            data["Exchange Rate"] = info[2]
       
        
        # Extract payment due
        match = re.search(r"Vencimento\s*\n?(\d{4}-\d{2}-\d{2})", text)
        if match:
            data["Payment Due"] = match.group(1)
        
        # Extract commercial discount
        match = re.search(r"Desconto Comercial\s*\n?([\d.,]+)", text)
        if match:
            data["Commercial Discount"] = match.group(1)

        # Extract payment terms and payment due
        match = re.search(r"Condição Pagamento\s*(.+?)\n", text)
        print(match.group(1))
        if match:
            payment_info = match.group(1).split(" ")
            print(payment_info)
            data["Payment Due"] = payment_info[2]
            payment_info_length = len(payment_info)
            if payment_info_length > 4: 
                data["Payment Terms"] = payment_info[3] + " " + payment_info[4] #Pronto Pagamento
                if payment_info_length > 5:
                    data["Payment Terms"] += " " + payment_info[5] # Fatura 5 dias
            else:
                data["Payment Terms"] = payment_info[3]
        
        # Extract execution location
        match = re.search(r"Local de Execução:\s*(.+?)(?:\n|$)", text)
        if match:
            data["Execution Location"] = match.group(1).strip()
        
        # Extract VAT
        match = re.search(r"IVA\s*([\d.,]+)", text)
        if match:
            data["VAT"] = match.group(1)
        
        # Extract customer name and address
        match = re.search(r"Exmo.\(s\) Sr.\(s\)\n(.+?)\n(.+?)\n(.+?)\n", text)
        if match:
            data["Customer"] = match.group(2).strip()
            data["Address"] = f"{match.group(3).strip()}"
        
        # Extract total 
        # Extract total amount
        match = re.search(r"Total \( EUR \)\s*([\d\s.,]+)", text)

        if match:
            data["Total Amount"] = match.group(1)

        
        # Extract total liquid
        match = re.search(r"IVA\s*\([\d.,]+\)\s*([\d\s.,]+)", text)
        #print(match.group(1))
        if match:
            data["Total Liquid"] = match.group(1).split(" ")[0]

        # Extract items
        item_pattern = re.compile(
            r"(7\d{7,}(?:\.\d{2})?)\s+"  # Artigo starting with 7
            r"(.+?)\s+"                 # Descrição
            r"(\d+,\d{2})\s+"           # Qtd.
            r"(\w+)"                    # Un.
            r"(?:\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+))?"  # Optional fields: Pr. Unitário, Desc., IVA, Valor  # Optional fields: Pr. Unitário, Desc., IVA, Valor

        )
        item_matches = item_pattern.findall(text)
        #print(f"Found {len(item_matches)} items in {pdf_path}")  # Debug print
        
        for item in item_matches:
            # print(item)
            data["Items"].append({
                "Invoice Number": data["Invoice Number"],
                "Date": data["Date"],
                "Currency": data["Currency"],
                #"Client_NIF": data["Client_NIF"],
                "Exchange Rate": data["Exchange Rate"],
                "Payment Due": data["Payment Due"],
                "Commercial Discount": data["Commercial Discount"],
                "Payment Terms": data["Payment Terms"],
                "Customer": data["Customer"],
                "Address": data["Address"],
                "Total Amount": data["Total Amount"],
                "Total Liquid": data["Total Liquid"],
                "Execution Location": data["Execution Location"],
                "VAT": data["VAT"],
                "Artigo": item[0],
                "Descrição": item[1],
                "Qtd.": item[2],
                "Un.": item[3],
                "Pr. Unitário": item[4],
                "Desc.": item[5] ,
                "Tax_Rate": item[6],
                "Valor Liq. Artigo": item[7]
            })
    except Exception as e:
        print(f"Failed to extract data from {pdf_path}: {e}")
    
    return data

# Process all PDF files in the directory
all_data = []
for pdf_path in glob.glob(pdf_path_pattern):
    data = extract_invoice_data(pdf_path)
    all_data.extend(data["Items"])

# Convert to DataFrame
df_items = pd.DataFrame(all_data)

# Create an Excel file
excel_path = "invoice_data.xlsx"
df_items.to_excel(excel_path, sheet_name="Invoice Data", index=False)

print(f"Excel file saved at {excel_path}")
