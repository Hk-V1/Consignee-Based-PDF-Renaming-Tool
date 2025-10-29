# Consignee-Based PDF Renaming Tool

## Summary
A lightweight **Python-based Windows desktop application** that automates the renaming of shipment or invoice PDFs using consignee information.  
The tool extracts each file’s **“Consignee (Ship To)”** name, cleans unwanted text (like *“Buyer’s Order No.”* or *“Dated”*), and exports the renamed files as a neatly packaged ZIP — ideal for logistics and documentation teams.

---

## Tech Stack
- **Language:** Python 3  
- **GUI Framework:** Tkinter  
- **Libraries:** `pdfplumber`, `zipfile`, `tempfile`, `shutil`, `re`, `os`, `uuid`, `pathlib`

---

## Key Features
- Upload a ZIP file containing multiple PDFs  
- Automatically extract and list all PDFs in the interface  
- Detect consignee names from **“Consignee (Ship To)”** or **“Ship To”** fields  
- Clean extracted names by removing text like *“Buyer’s Order No.”* or *“Dated”*  
- Rename files intelligently with serial numbers for duplicates  
- Export all renamed PDFs as a single downloadable ZIP  
- Simple Tkinter GUI — no coding required  

---

## Workflow
1. **Upload ZIP** → User uploads a ZIP containing PDF documents  
2. **Extract PDFs** → App lists all PDFs in the left panel  
3. **Rename Files** → App reads consignee name and renames each file  
4. **Export ZIP** → Download the final ZIP of renamed PDFs  

---

## Use Case
Perfect for **logistics, invoicing, and documentation teams** handling large batches of shipment or billing PDFs.  
This tool ensures **clean, consistent, and automated file naming** based on consignee details — saving hours of manual effort.

