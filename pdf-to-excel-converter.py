import pdfplumber
import pandas as pd
import re
import os


def extract_numbers_from_pdf(pdf_path):
    """extract numbers from a pdf file"""
    if not os.path.exists(pdf_path):
        print(f"File {pdf_path} does not exist")
        return None

    all_data = []

    with pdfplumber.open(pdf_path) as pdf:
        print(f"Number of pages: {len(pdf.pages)}\n")

        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if not text:
                continue

            numbers = re.findall(r"[\d,]+\.?\d*", text)
            for num in numbers:
                clean_num = num.replace(",", "")
                try:
                    value = float(clean_num)
                    all_data.append({
                        "page": page_num,
                        "raw": num,
                        "value": value
                    })
                except ValueError:
                    continue

    return all_data


def save_to_excel(data, output_path):
    """save data to excel"""
    if not data:
        print("No data")
        return

    df = pd.DataFrame(data)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='numbers', index=False)

        summary = pd.DataFrame({
            "Number of numbers": [len(df)],
            "total": [df["value"].sum()],
            "average": [df["value"].mean()],
            "max": [df["value"].max()],
            "min": [df["value"].min()]
        })
        summary.to_excel(writer, sheet_name='summary', index=False)

    print(f"saving file excel: {output_path}")
    print(f"number of numbers extracted: {len(df)}")
    print(f"total numbers: {df['value'].sum()}")


def main():
    print("Extract numbers from PDF to Excel\n")

    pdf_path = input("Enter path of PDF: ").strip()
    if not pdf_path:
        pdf_path = "sample_invoice.pdf"

    output_path = input("Enter path of Excel File: ").strip()
    if not output_path:
        output_path = "output.xlsx"

    print("\nExtracting...\n")
    data = extract_numbers_from_pdf(pdf_path)

    if data:
        save_to_excel(data, output_path)
        print("Done!\n")


if __name__ == "__main__":
    main()