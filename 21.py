import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
import threading

class ExcelEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Viewer")

        # Set the window size
        self.root.geometry("900x600")

        # Load the Excel file in a separate thread
        threading.Thread(target=self.load_excel_in_thread).start()

    def load_excel_in_thread(self):
        # File path provided here
        self.file_path = "C:/Users/Admin/Downloads/Spectrum Proposals 2024.xlsx"
        if self.file_path:
            self.load_excel()
            self.create_widgets()
        else:
            messagebox.showerror("Error", "No file selected!")
            self.root.quit()

    def load_excel(self):
        try:
            # Load the Excel file and drop unnamed or empty columns/rows
            self.df = pd.read_excel(self.file_path)
            
            # Drop completely empty columns and rows
            self.df.dropna(how='all', axis=1, inplace=True)  # Drop empty columns
            self.df.dropna(how='all', inplace=True)  # Drop empty rows

            print(self.df.head())  # Print the cleaned data to check
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {e}")
            self.root.quit()

    def create_widgets(self):
        self.root.after(0, self._create_widgets)

    def _create_widgets(self):
        # Define the columns you want to display in the Treeview
        columns = ['Date ', 'Agent Name', 'Business Name', 'Address', 'Offer', 'Price', 'Email']

        # Create a Treeview to display the cleaned DataFrame with scrollbars
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", height=15)

        # Set up columns and headings
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor='center')

        # Insert data into the Treeview
        for idx, row in self.df.iterrows():
            # Get the relevant columns from the DataFrame
            self.tree.insert("", "end", values=(
                row.get('Date ', ''), 
                row.get('Agent Name', ''), 
                row.get('Business Name', ''), 
                row.get('Address', ''), 
                row.get('Offer', ''), 
                row.get('Price', ''), 
                row.get('Email', '')
            ))

        # Add scrollbars for navigation
        vsb = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.root, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def show_order_details(self, event):
        # Get the selected row
        item = self.tree.selection()[0]
        row_values = self.tree.item(item, 'values')  # Get all values of the selected row

        # Create a new window to display detailed information
        self.detail_window = tk.Toplevel(self.root)
        self.detail_window.title("Order Details")
        self.detail_window.geometry("600x500")

        # Format the row's information as per your structure
        detailed_info = f"""
        -------------------------------------------------------
        Industry: NGO
        Business Name: {row_values[2]} 
        Address: {row_values[3]}
        City: COLUMBUS
        State: OH
        Zip: 43214

        First Name: APRIL 
        Last Name: HOY

        Main: 6144818227
        Additional No: N/A

        Cell No:  {row_values[6]}
        Email:  {row_values[5]}

        -------------------------------------------------------

        Sale Date: {row_values[0]}
        Agent Name: {row_values[1]}
        Sale closed by: {row_values[1]}
        Extension Number: N/A

        Current LEC: Spectrum
        Services Offered:  {row_values[4]}
        Price: {row_values[5]}

        -------------------------------------------------------

        Agent Comments: Cx is already using spectrum for 1 TN and internet .............. /////EMAIL OFFER : - 1 TN (Existing) + 1 TN  NEW + Internet 600 mbps with free business wifi bcoz of promotion =  $99.97 .........Cx need Installation  Between Mon - Fri  ......... 2nd POC name is KEITH GOODWIN  ....... CELL PHONE No - 740-817-2000 Email - info@campwyandot.org  .... NOTE : - (" After Installation we need to cancel old internet ......  Current Bill comes under April Hoy name .... We have 2nd POC as KEITH GOODWIN.")

        Mail to Sent cx?: {row_values[5]}
        Recording? Y/N: YES
        Notes: Chris: offer $69.98 (1 TN NEW + Internet 600 mbps with Free Business Wifi bcoz of Promotion - $69.98) price guarantee up to three years & 1 free domain also free of cost.
        cx: OK, go ahead   //  2nd poc - Keith Goodwin  // cx was good  // Remarks: g2g

        -------------------------------------------------------
        """

        # Display the detailed information in a text box
        text_box = tk.Text(self.detail_window, wrap=tk.WORD, width=80, height=25)
        text_box.insert(tk.END, detailed_info)
        text_box.config(state=tk.DISABLED)  # Make the text box read-only
        text_box.pack(padx=10, pady=10)

# Create the Tkinter app
root = tk.Tk()
app = ExcelEditor(root)
root.mainloop()
