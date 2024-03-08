import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import fitz  # PyMuPDF
import io

def open_pdf_page(pdf_path, page_num=0):
    """Open a PDF and return a PIL Image of the specified page."""
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_num)  # load the specified page
    pix = page.get_pixmap()  # render page to an image
    img = Image.open(io.BytesIO(pix.tobytes("ppm")))
    return img

def on_click(event, canvas):
    """Handle click events on the canvas."""
    x, y = event.x, event.y
    print(f"Clicked at x: {x}, y: {y}")
    # Here, assign these values to your configuration variables
    # Optionally, mark the click position on the canvas
    canvas.create_oval(x - 5, y - 5, x + 5, y + 5, fill='red')

def main():
    root = tk.Tk()
    root.title("Select Configuration Points")

    # Ask the user to select a PDF file
    pdf_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF Files", "*.pdf")])

    if not pdf_path:
        print("No file selected.")
        return  # Exit if no file is selected

    # Open the selected PDF and convert to a format Tkinter can use
    img = open_pdf_page(pdf_path)
    photo = ImageTk.PhotoImage(image=img)

    # Create a canvas to display the image
    canvas = tk.Canvas(root, width=img.width, height=img.height)
    canvas.pack()

    # Display the image
    canvas.create_image(0, 0, anchor=tk.NW, image=photo)

    # Bind the click event
    canvas.bind("<Button-1>", lambda event, canvas=canvas: on_click(event, canvas))

    root.mainloop()

if __name__ == "__main__":
    main()
