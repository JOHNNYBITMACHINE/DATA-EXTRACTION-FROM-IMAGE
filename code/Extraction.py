import google.generativeai as genai
import pandas as pd
from PIL import Image
from io import StringIO
import time
import openpyxl
import os
import warnings
import re
from datetime import datetime



warnings.filterwarnings("ignore")

class InvoiceExtractor:
    def __init__(self, api_key, image_folder, output_path):
        self.api_key = api_key
        self.image_folder = image_folder
        self.output_path = output_path
        genai.configure(api_key=self.api_key)
        self.prompts = [
            """
            This image contains tabular invoice data. Please extract the following columns:
            - Date
            - Seller Name
            - Seller GST
            - Seller Email ID
            - Seller Mobile No.
            - Buyer Name
            - Buyer GSTN
            - Buyer Mobile No.
            - Seller Invoice No
            - Discount
            If any value is missing, set it as NaN. Return the data as a CSV string with headers.
            """,
            """
            This image contains tabular invoice data. Please extract the following columns:
            -Date
            -Seller Invoice No
            -Seller GST
            -Item_Service Name
            -Batch
            -Expiary Date
            -Quantity/QTY
            -Unit Of Measurement
            -Rate
            -MRP
            -Amount
            If any value is missing, set it as NaN. Return the data as a CSV string with headers.
            """,
            """
            This image contains tabular invoice data. Please extract the following columns:
            -Date
            -Seller Invoice No
            -Seller GST	
            -HSN/SAC Code
            -Amount
            -Taxable amount
            -Central tax rate%
            -CGST amount
            -State tax rate%
            -SGST amount
            -Total tax amount
            If any value is missing, set it as NaN. Return the data as a CSV string with headers.
            """
        ]
        self.now = datetime.now()

    def extract_csv(self, text):
        csv_match = re.search(r"```csv\s*(.*?)\s*```", text, re.DOTALL)
        return csv_match.group(1) if csv_match else text

    def process_images(self):
        image_names = [f for f in os.listdir(self.image_folder) if os.path.isfile(os.path.join(self.image_folder, f))]
        print(self.now.strftime("%Y-%m-%d %H:%M:%S"))
        for i in image_names:
            image_path = os.path.join(self.image_folder, i)
            image = Image.open(image_path)
            model = genai.GenerativeModel("gemini-2.5-flash-preview-05-20")
            dfs = []
            for idx, prompt in enumerate(self.prompts):
                response = model.generate_content([prompt, image])
                text = response.text
                print(f"Gemini Output for Prompt {idx+1} (CSV):\n", text)
                csv_content = self.extract_csv(text)
                df = pd.read_csv(StringIO(csv_content), skip_blank_lines=True)
                df['updated_at'] = self.now
                dfs.append(df)
                time.sleep(25 if idx < 2 else 2)
            self.append_to_excel(dfs)
            print(f"Data appended row-wise to {self.output_path}")
            time.sleep(25)



    def append_to_excel(self, dfs):
        if os.path.exists(self.output_path):
            with pd.ExcelFile(self.output_path) as reader:
                try:
                    old_df1 = pd.read_excel(reader, sheet_name='Invoice Info')
                except:
                    old_df1 = pd.DataFrame()
                try:
                    old_df2 = pd.read_excel(reader, sheet_name='Item Details')
                except:
                    old_df2 = pd.DataFrame()
                try:
                    old_df3 = pd.read_excel(reader, sheet_name='Tax Details')
                except:
                    old_df3 = pd.DataFrame()
                    
        else:
            old_df1 = pd.DataFrame()
            old_df2 = pd.DataFrame()
            old_df3 = pd.DataFrame()
        df1_combined = pd.concat([old_df1, dfs[0]], axis=0, ignore_index=True)
        df2_combined = pd.concat([old_df2, dfs[1]], axis=0, ignore_index=True)
        df3_combined = pd.concat([old_df3, dfs[2]], axis=0, ignore_index=True)
        with pd.ExcelWriter(self.output_path, engine='openpyxl', mode='w') as writer:
            df1_combined.to_excel(writer, sheet_name='Invoice Info', index=False)
            df2_combined.to_excel(writer, sheet_name='Item Details', index=False)
            df3_combined.to_excel(writer, sheet_name='Tax Details', index=False)


extractor = InvoiceExtractor(
api_key="AIzaSyAd91qgrRX9CibBrG_DSnEYa2Y_EH7zCTQ",
image_folder="E:/PDL 2 - INVOICE/image",                          # image folder 
output_path="E:/PDL 2 - INVOICE/output/invoice_output.xlsx")
extractor.process_images()