import comtypes.client
import os

def create_pdf(path, ppt_file_name, pdf_file_name, formatType = 32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    # Ensure the path ends with a slash before concatenating
    if not path.endswith(os.sep):
        path += os.sep

    # Join paths properly
    ppt_path = path + ppt_file_name
    pdf_path = path + pdf_file_name
    
    # Replace forward slashes with backslashes for Windows
    ppt_path = ppt_path.replace("/", "\\")
    pdf_path = pdf_path.replace("/", "\\")

    print(f"PowerPoint path: {ppt_path}")
    print(f"PDF path: {pdf_path}")

    try:
        deck = powerpoint.Presentations.Open(ppt_path)
        deck.SaveAs(pdf_path, formatType)  # formatType = 32 for ppt to pdf
        deck.Close()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        powerpoint.Quit()