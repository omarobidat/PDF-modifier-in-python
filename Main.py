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
import platform
import threading

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

def is_pdf_file(filename):
    return filename.lower().endswith('.pdf')

directoryofpicking = ''
directoryofpacking = ''
cancel_flag = False
fixervaraible = 1
def check_file_access_with_retry(path):
    if not os.path.exists(path):
        print(f"The file {path} does not exist.")
        return False
    
    try:
        with open(path, 'a') as file_handle:
            if platform.system() == 'Windows':
                # On Windows, just being able to open the file in append mode is our check
                return True
            else:
                return False
    except:
        return False

def process_picking_directory():
    global directoryofpicking, cancel_flag
    while not cancel_flag:
        for filename in os.listdir(directoryofpicking):
            time.sleep(5)

            progress_bar.UpdateBar(0)
            input_pdf_path = os.path.join(directoryofpicking, filename)
            file_is_pdf = is_pdf_file(input_pdf_path)
            file_is_accessable = check_file_access_with_retry(input_pdf_path)
            if not file_is_pdf or not file_is_accessable:
                continue
            if file_is_pdf and file_is_accessable:
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
                        sg.popup_error(f'Error1: {str(e)}')
                try:
                        os.remove(input_pdf_path)
                        pass
                except Exception as e:
                    sg.popup_error(f'Error1: {str(e)}')


def process_packing_directory():
    global directoryofpacking, cancel_flag
    while not cancel_flag:
        for filename in os.listdir(directoryofpacking):
            time.sleep(5)

            progress_bar.UpdateBar(0)
            input_pdf_path = os.path.join(directoryofpacking, filename)
            file_is_pdf = is_pdf_file(input_pdf_path)
            file_is_accessable = check_file_access_with_retry(input_pdf_path)
            if not file_is_pdf or not file_is_accessable:
                continue
            
            if file_is_pdf and file_is_accessable:
                try:
                                window['-STATUS-'].update(f'Processing file: {filename}')
                                progress_bar.UpdateBar(500)
                                pdf_path = filename
                                input_pdf_path = os.path.join(directoryofpacking, filename)
                                doc = fitz.open(input_pdf_path)
                                total_pages = len(doc)
                                flipcolor = 1
                                gray = (230 / 255, 230 / 255, 230 / 255)
                                sumofallboxes = 0
                                for page_num in range(len(doc)):
                                    page = doc[page_num]
                                    start_y = 335
                                    line_height = 11.5058823529
                                    item_code_start_x = 473.5
                                    item_code_end_x = 525.6
                                    quantity_start_x = 527
                                    quantity_end_x = 574
                                    current_y = start_y
                                    midstart = 525.6
                                    midend = 527
                                    while current_y + line_height <= 700:
                                        rect_item_code = fitz.Rect(item_code_start_x, current_y, item_code_end_x, current_y + line_height)
                                        rect_quantity = fitz.Rect(quantity_start_x, current_y, quantity_end_x, current_y + line_height)
                                        item_code_text = page.get_text("text", clip=rect_item_code).strip()
                                        quantity_text = page.get_text("text", clip=rect_quantity).strip()

                                        midlane = fitz.Rect(midstart, current_y, midend, current_y + line_height)
                                        item_code_text = item_code_text.replace(",", "")

                                        if item_code_text and quantity_text:
                                            try:
                                                item_code = int(item_code_text)
                                                quantity = int(quantity_text)
                                                division_result = item_code / quantity if quantity != 0 else 0
                                                sumofallboxes = sumofallboxes + division_result
                                                if flipcolor == 1:
                                                    page.draw_rect(rect_item_code, color=(1, 1, 1), fill=(1, 1, 1))
                                                    page.draw_rect(rect_quantity, color=(1, 1, 1), fill=(1, 1, 1))
                                                    page.draw_rect(midlane, color=(1, 1, 1), fill=(1, 1, 1))
                                                    flipcolor = 0
                                                else:
                                                    page.draw_rect(rect_item_code, color=gray, fill=gray)
                                                    page.draw_rect(rect_quantity, color=gray, fill=gray)
                                                    page.draw_rect(midlane, color=gray, fill=gray)

                                                    flipcolor = 1
                                                page.insert_text((rect_item_code.x0+10, rect_item_code.y0+8), str(int(division_result)), color=(0, 0, 0), fontsize=10)
                                            except ValueError:
                                                pass
                                        elif item_code_text:
                                            if flipcolor == 1:
                                                    page.draw_rect(rect_item_code, color=(1, 1, 1), fill=(1, 1, 1))
                                                    page.draw_rect(rect_quantity, color=(1, 1, 1), fill=(1, 1, 1))
                                                    page.draw_rect(midlane, color=(1, 1, 1), fill=(1, 1, 1))

                                            else:
                                                    page.draw_rect(rect_item_code, color=gray, fill=gray)
                                                    page.draw_rect(rect_quantity, color=gray, fill=gray)
                                                    page.draw_rect(midlane, color=gray, fill=gray)
                                        else:
                                            if flipcolor == 0:
                                                    page.draw_rect(midlane, color=(1, 1, 1), fill=(1, 1, 1))
                                            else:
                                                    page.draw_rect(midlane, color=gray, fill=gray)
                                        current_y += line_height
                                    rect = fitz.Rect(473+0.5, 311+0.5, 575.5-0.5, 332.4-0.5)
                                    color = (230 / 255, 230 / 255, 230 / 255)
                                    page.draw_rect(rect, color=color, fill=color,)
                                    text = "Boxes"
                                    page.insert_textbox(rect, text, fontsize=11, align=1)
                                    rect123 = fitz.Rect(280, 751, 370, 770)
                                    page.draw_rect(rect123, color=(1, 1, 1), fill=(1, 1, 1))
                                    rect1234 = fitz.Rect(280, 751, 370, 770)

                                    page.draw_rect(rect1234, color=(1, 1, 1), fill=(1, 1, 1))

                                    page.insert_textbox(rect123, f'Page {page_num+1} of {total_pages}', fontsize=11, align=1)
                                finalpage = doc[len(doc)-1]
                                rect123 = fitz.Rect(400, 740, 480, 790)
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
        cancel_flag = True  # Set the flag to True to stop the threads
        break
    if event == '-FOLDER1-':
        directoryofpacking = values['-FOLDER1-']
        window['-STATUS-'].update(f'Monitoring directory: {directoryofpacking}')
        if directoryofpacking != '':
            packing_thread = threading.Thread(target=process_packing_directory)

            packing_thread.start()


    if event == '-FOLDER2-':
        directoryofpicking = values['-FOLDER2-']
        window['-STATUS-'].update(f'Monitoring directory: {directoryofpicking}')
        if directoryofpicking != '':
            picking_thread = threading.Thread(target=process_picking_directory)
            picking_thread.start()

window.close()
