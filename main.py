import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
from pdf2image import convert_from_path
import PyPDF2


class PDFEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Rearranger and Rotator")
        self.root.geometry("1520x1000")

        # Variables to hold the PDF information
        self.pdf_path = None  # Original PDF file path
        self.pdf_images = []  # List of PIL images (one per PDF page)
        # List of dictionaries: each entry has the original page index, its number, and rotation value.
        self.page_order = []
        self.preview_photo = None  # To hold a reference to the preview PhotoImage

        # Build the UI components
        self.create_widgets()

    def create_widgets(self):
        # Top frame for the "Load PDF" button
        top_frame = ttk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        load_btn = ttk.Button(top_frame, text="Load PDF", command=self.load_pdf)
        load_btn.pack(side=tk.LEFT)

        # Main frame: left side for the rearrangement table; right side for the preview pane.
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

        # -------------------------
        # Left Frame – Rearrangement Table & Controls
        # -------------------------
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Create a Treeview with three columns: Serial, Current Page, Rotation (°)
        self.tree = ttk.Treeview(left_frame, columns=("serial", "current_page", "rotation"), show="headings",
                                 selectmode="browse", height=20)
        self.tree.heading("serial", text="Serial")
        self.tree.heading("current_page", text="Current Page")
        self.tree.heading("rotation", text="Rotation (°)")
        self.tree.column("serial", width=50, anchor="center")
        self.tree.column("current_page", width=100, anchor="center")
        self.tree.column("rotation", width=100, anchor="center")
        self.tree.pack(side=tk.TOP, fill=tk.Y, padx=5, pady=5)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        # Controls for reordering and updating rotation
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        up_btn = ttk.Button(btn_frame, text="Move Up", command=self.move_up)
        up_btn.pack(side=tk.LEFT, padx=2, pady=2)

        down_btn = ttk.Button(btn_frame, text="Move Down", command=self.move_down)
        down_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # Entry to update rotation for the selected row
        self.rotation_entry = ttk.Entry(btn_frame, width=5)
        self.rotation_entry.pack(side=tk.LEFT, padx=5)
        update_rot_btn = ttk.Button(btn_frame, text="Set Rotation", command=self.update_rotation)
        update_rot_btn.pack(side=tk.LEFT, padx=2)

        # Button to export the modified PDF
        export_btn = ttk.Button(left_frame, text="Export PDF", command=self.export_pdf)
        export_btn.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        # -------------------------
        # Right Frame – PDF Preview Pane
        # -------------------------
        right_frame = ttk.Frame(main_frame, relief=tk.SUNKEN, borderwidth=2, style="Black.TFrame")
        right_frame.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH, padx=5, pady=5)
        self.preview_label = ttk.Label(right_frame, background="black")
        self.preview_label.pack(expand=True, fill=tk.BOTH)
        self.preview_label.bind("<Configure>", self.update_preview)

    def load_pdf(self):
        """Load a PDF file, convert its pages to images, and populate the table."""
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not file_path:
            return

        self.pdf_path = file_path

        try:
            # Convert PDF pages to images
            self.pdf_images = convert_from_path(self.pdf_path, poppler_path=r"C:\Program Files\poppler\Library\bin", dpi=400)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert PDF to images:\n{e}")
            return

        # Initialize page_order: each page has its original index, page number, and default rotation of 0.
        self.page_order = []
        for i in range(len(self.pdf_images)):
            self.page_order.append({
                "orig_index": i,
                "page_number": i + 1,
                "rotation": 0
            })

        self.populate_treeview()
        messagebox.showinfo("Info", f"Loaded PDF with {len(self.pdf_images)} pages.")

    def populate_treeview(self):
        """Refresh the rearrangement table based on the current page_order list."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, page in enumerate(self.page_order):
            self.tree.insert("", "end", iid=str(idx), values=(idx + 1, page["page_number"], page["rotation"]))

    def on_tree_select(self, event):
        """When a row is selected, update the preview pane."""
        self.update_preview()

    def update_preview(self, event=None):
        """Display the preview of the selected page without cropping.

        If the image is portrait (or square), scale it to the container's height.
        If the image is landscape, scale it to the container's width.
        """
        # Get the selected page from the rearrangement table
        selected = self.tree.selection()
        if not selected:
            return

        idx = int(selected[0])
        page_info = self.page_order[idx]
        orig_index = page_info["orig_index"]
        rotation = page_info["rotation"]

        # Get the original image and apply rotation if needed.
        img = self.pdf_images[orig_index]
        if rotation != 0:
            rotated_img = img.rotate(-rotation, expand=True)
        else:
            rotated_img = img

        # Get the dimensions of the preview container.
        if event:
            container_width = event.width
            container_height = event.height
        else:
            container_width = self.preview_label.winfo_width()
            container_height = self.preview_label.winfo_height()

        # Fallback values if the widget isn't fully rendered yet.
        if container_width < 50:
            container_width = 500
        if container_height < 50:
            container_height = 700

        # Get the image dimensions.
        img_width, img_height = rotated_img.size

        # Determine the scaling factor based on the image's orientation.
        if img_height >= img_width:
            # Portrait (or square): scale to fill the container's height.
            scale = container_height / img_height
            new_width = int(img_width * scale)
            new_height = container_height
        else:
            # Landscape: scale to fill the container's width.
            scale = container_width / img_width
            new_width = container_width
            new_height = int(img_height * scale)

        new_size = (new_width, new_height)

        # Use high-quality resampling for resizing.
        try:
            resample = Image.Resampling.LANCZOS
        except AttributeError:
            resample = Image.ANTIALIAS

        resized_img = rotated_img.resize(new_size, resample)

        # Convert the PIL image to a Tkinter PhotoImage.
        self.preview_photo = ImageTk.PhotoImage(resized_img)
        self.preview_label.config(image=self.preview_photo)

    def move_up(self):
        """Move the selected row one position up."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "No row selected.")
            return
        idx = int(selected[0])
        if idx <= 0:
            return
        # Swap the current row with the previous row in page_order.
        self.page_order[idx], self.page_order[idx - 1] = self.page_order[idx - 1], self.page_order[idx]
        self.populate_treeview()
        self.tree.selection_set(str(idx - 1))
        self.update_preview()

    def move_down(self):
        """Move the selected row one position down."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "No row selected.")
            return
        idx = int(selected[0])
        if idx >= len(self.page_order) - 1:
            return
        self.page_order[idx], self.page_order[idx + 1] = self.page_order[idx + 1], self.page_order[idx]
        self.populate_treeview()
        self.tree.selection_set(str(idx + 1))
        self.update_preview()

    def update_rotation(self):
        """Update the rotation value for the selected page."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "No row selected.")
            return
        try:
            rotation_val = int(self.rotation_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid integer for rotation.")
            return

        # Optionally warn if the rotation is not a multiple of 90.
        if rotation_val % 90 != 0:
            if not messagebox.askyesno("Warning", "Rotation is not a multiple of 90. Continue?"):
                return

        idx = int(selected[0])
        self.page_order[idx]["rotation"] = rotation_val
        self.populate_treeview()
        self.tree.selection_set(str(idx))
        self.update_preview()

    def export_pdf(self):
        """Export the rearranged and rotated PDF using PyPDF2."""
        if not self.pdf_path:
            messagebox.showwarning("Warning", "No PDF loaded.")
            return

        try:
            with open(self.pdf_path, "rb") as infile:
                reader = PyPDF2.PdfReader(infile)
                writer = PyPDF2.PdfWriter()

                # Iterate over pages in the new order
                for page_info in self.page_order:
                    orig_index = page_info["orig_index"]
                    rotation = page_info["rotation"]

                    page = reader.pages[orig_index]
                    if rotation != 0:
                        # For positive rotation values, use rotate_clockwise.
                        if rotation > 0:
                            page.rotate(rotation)
                        else:
                            # For negative values, rotate clockwise by (360 + rotation)
                            page.rotate(360 + rotation)
                    writer.add_page(page)

                # Ask the user where to save the new PDF.
                output_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF Files", "*.pdf")],
                    title="Save Rearranged PDF As"
                )
                if not output_path:
                    return

                with open(output_path, "wb") as outfile:
                    writer.write(outfile)

            messagebox.showinfo("Success", "PDF exported successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while exporting the PDF:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFEditor(root)
    root.mainloop()
