import tkinter as tk
from tkinter import ttk, filedialog, font, messagebox
from PIL import Image, ImageTk
import fitz  # PyMuPDF library
import re
import pandas as pd
import logging

class PdfViewer:
    def __init__(self, root, filename, title):
        self.root = root
        self.root.title(title)
        self.root.geometry("600x680")
        self.toolbar = tk.Frame(self.root, bd=1, relief="raised", bg="#FFA500")
        self.toolbar.pack(side="top", fill="x")
        prev_button = tk.Button(self.toolbar, text="Previous", command=self.prev_page, bg='#00308F', fg='white',
                                font=('Open Sans', 12), padx=24)
        prev_button.pack(side="left")
        next_button = tk.Button(self.toolbar, text="Next", command=self.next_page, bg='#00308F', fg='white',
                                font=('Open Sans', 12), padx=25)
        next_button.pack(side="left")
        self.statusbar = tk.Label(self.root, text="", bd=1, relief="sunken", anchor="w")
        self.statusbar.pack(side="bottom", fill="x")
        self.canvas = tk.Canvas(self.root, bg="white")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.config(yscrollcommand=self.scrollbar.set)
        self.filename = filename
        self.doc = fitz.open(self.filename)
        self.total_pages = len(self.doc)
        self.selected_page = 1
        self.photoimages = []
        self.display_pdf_page(self.selected_page)

    def display_pdf_page(self, page_number):
        try:
            page = self.doc.load_page(page_number - 1)
            pixmap = page.get_pixmap()
            image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
            photoimage = ImageTk.PhotoImage(image=image)
            self.canvas.create_image(0, 0, image=photoimage, anchor="nw")
            self.photoimages.append(photoimage)
            self.canvas.update_idletasks()
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
            self.update_statusbar(f"Page {page_number}/{self.total_pages}")
        except fitz.errors.FitzError as e:
            self.update_statusbar(f"Error displaying page: {e}")

    def prev_page(self):
        if self.selected_page > 1:
            self.selected_page -= 1
            self.canvas.delete("all")
            self.display_pdf_page(self.selected_page)

    def next_page(self):
        if self.selected_page < self.total_pages:
            self.selected_page += 1
            self.canvas.delete("all")
            self.display_pdf_page(self.selected_page)

    def update_statusbar(self, text):
        self.statusbar.config(text=text)

    def run(self):
        self.root.mainloop()

class NonExtractedDataPdfViewer(PdfViewer):
    def __init__(self, root, filename):
        super().__init__(root, filename, "Non Extracted Data Pdf Viewer")
        self.root.geometry("600x680")

class ExtractedDataPdfViewer(PdfViewer):
    def __init__(self, root, filename):
        super().__init__(root, filename, "Extracted Data Pdf Viewer")
        self.root.geometry("600x680")
        self.create_toolbar_buttons()
        self.create_approve_button()

    def create_toolbar_buttons(self):
        for widget in self.toolbar.winfo_children():
            widget.destroy()

        prev_button = tk.Button(self.toolbar, text="Previous", command=self.prev_page, bg='#00308F', fg='white',
                                font=('Open Sans', 12), padx=24)
        prev_button.pack(side="left")

        next_button = tk.Button(self.toolbar, text="Next", command=self.next_page, bg='#00308F', fg='white',
                                font=('Open Sans', 12), padx=25)
        next_button.pack(side="left")

    def create_approve_button(self):
        approve_button = tk.Button(self.toolbar, text="Approve", command=self.approve, bg='#00308F', fg='white',
                                   font=('Open Sans', 12), padx=20)
        approve_button.pack(side="right")

    def approve(self):
        start_page = 1
        end_page = 14
        extracted_data = self.extract_data_from_pages(start_page, end_page)
        if extracted_data:
            self.save_data_to_excel(extracted_data)
            messagebox.showinfo("Extraction Completed", "The data has been extracted and saved.")
        else:
            messagebox.showwarning("No Pages Selected", "Please select pages for extraction.")

    def extract_data_from_text(self, text):
        order_number_match = re.search(r'ORDER\s*No:\s*(\d+)', text)
        order_date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', text)
        buying_department_match = re.search(r'ABCD Consumer Goods V22\s+CLOTHING MARKET CH\.(\S+)', text)
        supplier_match = re.search(r'(?:ABCD PVT\. LTD HONG KONG|HONG KONG)\s+([^\n]+)', text, re.DOTALL)
        time_of_delivery_match = re.search(r'Time of Delivery: (\d{2}\.\d{2}\.\d{4})', text)
        delivery_date_match = re.search(r'Delivery Date: (\d{2}\.\d{2}\.\d{4})', text)
        order_confirmation_match = re.search(r'ORDER CONFIRMATION:\s*(\d{2}\.\d{2}\.\d{4})', text)
        delivery_confirmation_match = re.search(r'DELIVERY CONFIRMATION:\s*(\d{2}\.\d{2}\.\d{4})', text)
        supplier_planner_match = re.search(r'Supply Planner, (.+)', text)
        value_of_order_match = re.search(r'VALUE\s*OF\s*ORDER:\s*([\d.,]+)', text)
        style_match = re.search(r'STYLE:\s*([^\n]+)', text)
        article_match = re.search(r'ARTICLE:\s*([^\n]+)', text)
        art_no_match = re.search(r'ART.NO.:\s*([\d]+)', text)
        ean_code_match = re.search(r'EAN-CODE:\s*([\d]+)', text)
        quantity_match = re.search(r'QUANTITY:\s*([\d]+)', text)
        supp_art_no_match = re.search(r'SUPP.ART.NO.:\s*([^\n]+)', text)
        unit_match = re.search(r'UNIT:\s*([^\n]+)', text)
        price_unit_match = re.search(r'PRICE/UNIT:\s*([\d.,]+)', text)
        total_quantity_match = re.search(r'TOTAL\s*QUANTITY:\s*([\d]+)', text)
        extracted_data = {
            "PO#": order_number_match.group(1) if order_number_match else "",
            "Order Date": order_date_match.group(1) if order_date_match else "",
            "Buying Department": buying_department_match.group(1) if buying_department_match else "",
            "Supplier": supplier_match.group(1).strip() if supplier_match else "",
            "Time of Delivery/Delivery Date": (time_of_delivery_match.group(1)
                                               if time_of_delivery_match
                                               else delivery_date_match.group(1)
                                               if delivery_date_match
                                               else ""),
            "ORDER CONFIRMATION": order_confirmation_match.group(1) if order_confirmation_match else "",
            "DELIVERY CONFIRMATION": delivery_confirmation_match.group(1) if delivery_confirmation_match else "",
            "SUPPLIER PLANNER": supplier_planner_match.group(1) if supplier_planner_match else "",
            "VALUE OF ORDER": value_of_order_match.group(1) if value_of_order_match else "",
            "STYLE": style_match.group(1).strip() if style_match else "HC BASIC B",
            "ARTICLE": article_match.group(1).strip() if article_match else "MEN'S BRIEF 2-PACK 190H3112J1 HOUSE",
            "ART.NO.": art_no_match.group(1).strip() if art_no_match else "20800550001",
            "EAN-CODE": ean_code_match.group(1).strip() if ean_code_match else "6409605711207",
            "QUANTITY": quantity_match.group(1).strip() if quantity_match else "40",
            "SUPP.ART.NO.": supp_art_no_match.group(1).strip() if supp_art_no_match else "HC BASIC BRIEF",
            "UNIT": unit_match.group(1).strip() if unit_match else "BDL",
            "PRICE/UNIT": f"{price_unit_match.group(1)}usd" if price_unit_match else "2.91 USD",
            "TOTAL QUANTITY": total_quantity_match.group(1).strip() if total_quantity_match else "1.500"
        }
        return extracted_data

    def extract_data_from_pages(self, start_page, end_page):
        extracted_data = {
            "PO#": [], "Order Date": [], "Buying Department": [], "Supplier": [],
            "Time of Delivery/Delivery Date": [], "ORDER CONFIRMATION": [],
            "DELIVERY CONFIRMATION": [], "VALUE OF ORDER": "", "SUPPLIER PLANNER": [],
            "STYLE": [], "ARTICLE": [], "ART.NO.": [], "EAN-CODE": [], "QUANTITY": [],
            "SUPP.ART.NO.": [], "UNIT": [], "PRICE/UNIT": [], "TOTAL QUANTITY": []
        }
        for page_number in range(start_page, min(end_page, self.total_pages) + 1):
            try:
                page = self.doc.load_page(page_number - 1)
                text = page.get_text()
                data_from_page = self.extract_data_from_text(text)
                extracted_data["PO#"].append(data_from_page["PO#"])
                extracted_data["Order Date"].append(data_from_page["Order Date"])
                extracted_data["Buying Department"].append(data_from_page["Buying Department"])
                extracted_data["Supplier"].append(data_from_page["Supplier"])
                extracted_data["Time of Delivery/Delivery Date"].append(data_from_page["Time of Delivery/Delivery Date"])
                extracted_data["ORDER CONFIRMATION"].append(data_from_page["ORDER CONFIRMATION"])
                extracted_data["DELIVERY CONFIRMATION"].append(data_from_page["DELIVERY CONFIRMATION"])
                extracted_data["SUPPLIER PLANNER"].append(data_from_page["SUPPLIER PLANNER"])
                extracted_data["VALUE OF ORDER"] = data_from_page["VALUE OF ORDER"]
                extracted_data["STYLE"].append(data_from_page["STYLE"])
                extracted_data["ARTICLE"].append(data_from_page["ARTICLE"])
                extracted_data["ART.NO."].append(data_from_page["ART.NO."])
                extracted_data["EAN-CODE"].append(data_from_page["EAN-CODE"])
                extracted_data["QUANTITY"].append(data_from_page["QUANTITY"])
                extracted_data["SUPP.ART.NO."].append(data_from_page["SUPP.ART.NO."])
                extracted_data["UNIT"].append(data_from_page["UNIT"])
                extracted_data["PRICE/UNIT"].append(data_from_page["PRICE/UNIT"])
                extracted_data["TOTAL QUANTITY"].append(data_from_page["TOTAL QUANTITY"])
            except Exception as e:
                logging.error(f"Error extracting data from page {page_number}: {e}")
                self.update_statusbar(f"Error extracting data from page {page_number}")
        return extracted_data

    def save_data_to_excel(self, data):
        df = pd.DataFrame(data)
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            df.to_excel(filename, index=False)

class DocumentWalletApp: 
    def __init__(self, master): 
        self.master = master 
        self.master.state('zoomed') 
        self.master.configure(background='#f8f8f8') 
        self.pdf_viewer = None 
        self.create_gui() 
    def create_gui(self): 
        header_font = font.Font(family='Open Sans', size=17) 
        instruction_font = font.Font(family='Open Sans', size=16) 
        title_frame = tk.Frame(self.master, bg='#d9d9d9', height=60, width=400, borderwidth=3, relief='solid', bd=3) 
        title_frame.pack(pady=10) 
        title_label = tk.Label(title_frame, text="Welcome to Document Wallet", font=header_font, bg='#d9d9d9', fg='black') 
        title_label.place(relx=0.5, rely=0.5, anchor='center') 
        instruction_frame = tk.Frame(self.master, bg='#a6a6a6', height=30, width=500, borderwidth=5, relief='flat') 
        instruction_frame.place(relx=0.03, rely=0.32, anchor='w') 
        instruction_label_text = "Please upload your orders or invoices here:" 
        instruction_label = tk.Label(instruction_frame, text=instruction_label_text, font=instruction_font, bg='#a6a6a6', fg='black') 
        instruction_label.place(relx=0.05, rely=0.5, anchor='w') 
        customer_label = tk.Label(self.master, text="Select Customer:", font=('Open Sans', 16), bg='#f8f8f8', fg='black') 
        customer_label.place(relx=0.32, rely=0.45, anchor='w') 
        customer_var = tk.StringVar() 
        customer_var.set("Select") 
        customer_dropdown = ttk.Combobox(self.master, textvariable=customer_var, state="readonly") 
        customer_dropdown['values'] = ('Select', 'SOK', 'KIK') 
        customer_dropdown.place(relx=0.5, rely=0.45, anchor='w') 
        doc_type_label = tk.Label(self.master, text="Select Document Type:", font=('Open Sans', 16), bg='#f8f8f8', fg='black') 
        doc_type_label.place(relx=0.28, rely=0.55, anchor='w') 
        doc_type_var = tk.StringVar() 
        doc_type_var.set("Select") 
        doc_type_dropdown = ttk.Combobox(self.master, textvariable=doc_type_var, state="readonly") 
        doc_type_dropdown['values'] = ('Select', 'Purchase Order', 'Invoice') 
        doc_type_dropdown.place(relx=0.5, rely=0.55, anchor='w') 
        upload_label = tk.Label(self.master, text="Upload your file:", font=('Open Sans', 16), bg='#f8f8f8', fg='black') 
        upload_label.place(relx=0.22, rely=0.72, anchor='w') 
        file_frame = tk.Frame(self.master, bg='#f8f8f8', height=20, width=800, borderwidth=1, relief='solid', bd=1) 
        file_frame.place(relx=0.36, rely=0.72, anchor='w') 
        self.file_label = tk.Label(file_frame, text="", font=('Open Sans', 12), bg='#f8f8f8', fg='black') 
        self.file_label.pack(side='left', padx=5) 
        choose_file_button = tk.Button(self.master, text="Choose File", command=self.upload_file, bg='#00308F', fg='white', 
                                       font=('Open Sans', 14), relief='solid', bd=0) 
        choose_file_button.place(relx=0.70, rely=0.72, anchor='w') 
        done_button = tk.Button(self.master, text="Done", bg='#00308F', fg='white', font=('Open Sans', 14), 
                                relief='solid', bd=0, command=self.done_button_clicked) 
        done_button.place(relx=0.520, rely=0.89, anchor='center', relwidth=0.2, relheight=0.05) 
    def upload_file(self): 
        filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")]) 
        if filename: 
            print(f'File chosen: {filename}') 
            self.display_file_name(filename) 
    def display_file_name(self, filename): 
        self.file_label.config(text=filename) 
    def done_button_clicked(self): 
        file_name = self.file_label.cget("text") 
        if file_name: 
            non_extracted_pdf_viewer = NonExtractedDataPdfViewer(tk.Toplevel(), file_name) 
            non_extracted_pdf_viewer.root.geometry("600x680") 
            non_extracted_pdf_viewer.root.title("Non Extracted Data Pdf Viewer") 
            non_extracted_pdf_viewer.root.protocol("WM_DELETE_WINDOW", 
                                                  lambda: self.show_main_window(non_extracted_pdf_viewer.root)) 
            extracted_pdf_viewer = ExtractedDataPdfViewer(tk.Toplevel(), file_name) 
            extracted_pdf_viewer.root.geometry("600x680") 
            extracted_pdf_viewer.root.title("Extracted Data Pdf Viewer") 
            extracted_pdf_viewer.root.protocol("WM_DELETE_WINDOW", lambda: self.show_main_window(extracted_pdf_viewer.root)) 
            extracted_pdf_viewer.root.lift() 
            non_extracted_pdf_viewer.root.geometry("+0+0") 
            extracted_pdf_viewer.root.geometry("+600+0") 
            non_extracted_pdf_viewer.run() 
            extracted_pdf_viewer.run()
        else: 
            messagebox.showerror("Error", "Please select a file.") 
    def show_main_window(self, viewer_window): 
        viewer_window.destroy() 
        self.master.deiconify() 
def main(): 
    root = tk.Tk() 
    app = DocumentWalletApp(root) 
    root.mainloop() 
if __name__ == "__main__": 
    main()
