import os
import time
import pdfplumber
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import io
from PyPDF2 import PdfReader, PdfWriter
import fitz
import re
import PySimpleGUI as sg

layout = [
    [sg.Text('Monitoring Directory for packing slip:'), sg.InputText(key='-FOLDER1-', enable_events=True, size=(25, 1)), sg.FolderBrowse(target='-FOLDER1-')],
    [sg.Text('Monitoring Directory for picklist:'), sg.InputText(key='-FOLDER2-', enable_events=True, size=(25, 1)), sg.FolderBrowse(target='-FOLDER2-')],
    [sg.Text(size=(50, 2), key='-STATUS-')],
    [sg.ProgressBar(1000, orientation='h', size=(20, 20), key='-PROGBAR-')],
    [sg.Cancel()]
]

window = sg.Window('PDF Processor Status', layout)
progress_bar = window['-PROGBAR-']

def ensure_output_directory_exists_in_parent(directory):
    parent_dir = os.path.dirname(directory)
    return parent_dir

directoryofpicking = ""
directoryofpacking = ""
def update_pdf(pdf_path, output_pdf_path):
    try:
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        flipcolor = 1
        gray = (230 / 255, 230 / 255, 230 / 255)
        sumofallboxes = 0
        for page_num in range(len(doc)):
            page = doc[page_num]
            start_y = 336
            line_height = 13.8
            item_code_start_x = 390+20
            item_code_end_x = 480
            quantity_start_x = 480
            quantity_end_x = 570-20
            current_y = start_y
            while current_y + line_height <= 700:
                rect_item_code = fitz.Rect(item_code_start_x, current_y, item_code_end_x, current_y + line_height)
                rect_quantity = fitz.Rect(quantity_start_x, current_y, quantity_end_x, current_y + line_height)
                item_code_text = page.get_text("text", clip=rect_item_code).strip()
                quantity_text = page.get_text("text", clip=rect_quantity).strip()
                if item_code_text and quantity_text:
                    try:
                        item_code = int(item_code_text)
                        quantity = int(quantity_text)
                        division_result = item_code / quantity if quantity != 0 else 0
                        sumofallboxes = sumofallboxes + division_result
                        if flipcolor == 1:
                            page.draw_rect(rect_item_code, color=(1, 1, 1), fill=(1, 1, 1))
                            page.draw_rect(rect_quantity, color=(1, 1, 1), fill=(1, 1, 1))
                            flipcolor = 0
                        else:
                            page.draw_rect(rect_item_code, color=gray, fill=gray)
                            page.draw_rect(rect_quantity, color=gray, fill=gray)
                            flipcolor = 1
                        page.insert_text((rect_item_code.x0+10, rect_item_code.y0+10), str(division_result), color=(0, 0, 0), fontsize=11)
                    except ValueError:
                        pass
                elif item_code_text:
                    if flipcolor == 1:
                            page.draw_rect(rect_item_code, color=(1, 1, 1), fill=(1, 1, 1))
                            page.draw_rect(rect_quantity, color=(1, 1, 1), fill=(1, 1, 1))
                    else:
                            page.draw_rect(rect_item_code, color=gray, fill=gray)
                            page.draw_rect(rect_quantity, color=gray, fill=gray)
                current_y += line_height
            rect = fitz.Rect(390+3, 311+2, 570-3, 334-2)
            color = (230 / 255, 230 / 255, 230 / 255)
            page.draw_rect(rect, color=color, fill=color)
            text = "Boxes"
            page.insert_textbox(rect, text, fontsize=11, align=1)
            rect123 = fitz.Rect(280, 751, 370, 770)
            page.draw_rect(rect123, color=(1, 1, 1), fill=(1, 1, 1))
            page.insert_textbox(rect123, f'Page {page_num+1} of {total_pages}', fontsize=11, align=1)
        finalpage = doc[len(doc)-1]
        rect123 = fitz.Rect(400, 700, 480, 770)
        finalpage.draw_rect(rect123, color=(1, 1, 1), fill=(1, 1, 1))
        output_directory = ensure_output_directory_exists_in_parent(directoryofpacking)
        rect124 = fitz.Rect(504, 91, 575, 112)
        ordernumber = page.get_text("text", clip=rect124)
        ordernumber = ordernumber.split('\n')
        ext = ".pdf"
        new_base = "ps"+f'{ordernumber[0]}'
        finalpage.insert_textbox(rect123, f'Total number of Boxes {sumofallboxes}', fontsize=11, align=1)
        output_pdf_path = os.path.join(output_directory, new_base + ext)
        doc.save(output_pdf_path )
        doc.close()
        try:
                        os.remove(input_pdf_path)
        except PermissionError as e:
                    sg.popup_error(f"Failed to delete the file: {e}")
        progress_bar.UpdateBar(1000)
        window['-STATUS-'].update('Processing complete')
    except PermissionError as e:
        sg.popup_error(f"Failed to delete the file: {e}")

while True:
    event, values = window.read(timeout=100)
    if event in (sg.WIN_CLOSED, 'Cancel'):
        break
    if event == '-FOLDER1-':
        directoryofpacking = values['-FOLDER1-']
        window['-STATUS-'].update(f'Monitoring directory: {directoryofpacking}')
    if event == '-FOLDER2-':
        directoryofpicking = values['-FOLDER2-']
        window['-STATUS-'].update(f'Monitoring directory: {directoryofpicking}')
    if directoryofpicking:
        for filename in os.listdir(directoryofpicking):
            progress_bar.UpdateBar(0)
            try:
                window['-STATUS-'].update(f'Processing file: {filename}')
                progress_bar.UpdateBar(500)
                input_pdf_path = os.path.join(directoryofpicking, filename)
                with pdfplumber.open(input_pdf_path) as pdf, open(input_pdf_path, 'rb') as f:
                    original_pdf_reader = PdfReader(f)
                    total_pages = len(pdf.pages)
                    writer = PdfWriter()
                    totalnumberofboxes = 0
                    for i, page in enumerate(pdf.pages):
                        doc = fitz.open(input_pdf_path)
                        page1 = doc.load_page(i)
                        rect = fitz.Rect(152, 295, 215, 750)
                        text = page1.get_text("text", clip=rect)
                        rect = fitz.Rect(504, 90, 575, 112)
                        ordernumber = page1.get_text("text", clip=rect)
                        ordernumber = ordernumber.split('\n')
                        lines = text.split('\n')
                        try:
                            start_index = lines.index('Boxes') + 1
                        except ValueError:
                            start_index = len(lines)
                        total_boxes = sum(int(line) for line in lines[start_index:] if line.isdigit())
                        packet = io.BytesIO()
                        can = canvas.Canvas(packet, pagesize=letter)
                        totalnumberofboxes = total_boxes + totalnumberofboxes
                        if total_pages == i+1:
                            can.drawString(152, 30, f'Total Boxes: {totalnumberofboxes}')
                        can.setFillColorRGB(1, 1, 1)
                        can.rect(290, 25, 100, 15, fill=1)
                        can.setFillColorRGB(0, 0, 0)
                        can.drawString(290, 30, f'  Page {i+1} of {total_pages}')
                        can.save()
                        packet.seek(0)
                        new_pdf = PdfReader(packet)
                        overlay_page = new_pdf.pages[0]
                        existing_page = original_pdf_reader.pages[i]
                        existing_page.merge_page(overlay_page)
                        writer.add_page(existing_page)
                    new_base = "pl"+f'{ordernumber[0]}'
                    ext = '.pdf'
                    output_directory = ensure_output_directory_exists_in_parent(directoryofpicking)
                    output_pdf_path = os.path.join(output_directory, new_base + ext)
                    with open(output_pdf_path, 'wb') as f_out:
                        writer.write(f_out)
                    doc.close()
            except Exception as e:
                sg.popup_error(f'Error: {str(e)}')
            try:
                os.remove(input_pdf_path)
            except PermissionError as e:
                sg.popup_error(f"Failed to delete the file: {e}")    
    progress_bar.UpdateBar(1000)
    window['-STATUS-'].update('Processing complete')
    if directoryofpacking:
        for filename in os.listdir(directoryofpacking):
            progress_bar.UpdateBar(0)
            window['-STATUS-'].update(f'Processing file: {filename}')
            progress_bar.UpdateBar(500)
            pdf_path = filename
            input_pdf_path = os.path.join(directoryofpacking, filename)
            try:
                base, ext = os.path.splitext(filename)
                new_base = re.sub('packing list', '', base, flags=re.IGNORECASE) + "_fixed"
                output_directory = ensure_output_directory_exists_in_parent(directoryofpacking)
                output_pdf_path = os.path.join(output_directory, new_base + ext)
                update_pdf(directoryofpacking+"/"+pdf_path, output_pdf_path)
            except Exception as e:
                sg.popup_error(f'Error: {str(e)}')
    time.sleep(1)
window.close()
