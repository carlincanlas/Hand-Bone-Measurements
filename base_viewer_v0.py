import tkinter as tk
from tkinter import filedialog
import os
import pydicom
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class DICOMViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Hand DICOM Viewer")

        self.frame_index = 0
        self.num_frames = 1
        self.dicom = None
        self.pixel_data = None
        self.pixel_spacing = [1.0, 1.0]
        self.resize_after_id = None

        file_frame = tk.Frame(root)
        file_frame.pack(pady=5)

        self.filename_label = tk.Label(file_frame, text="No file loaded")
        self.filename_label.pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(file_frame, text="Load DICOM File", command=self.load_file).pack(side=tk.LEFT)

        self.label = tk.Label(root, text="")
        self.label.pack()

        nav_frame = tk.Frame(root)
        nav_frame.pack()

        tk.Button(nav_frame, text="Previous", command=self.prev_frame).pack(side=tk.LEFT, padx=5)
        tk.Button(nav_frame, text="Next", command=self.next_frame).pack(side=tk.LEFT, padx=5)

        self.frame_label = tk.Label(nav_frame, text="Frame 1 / 1")
        self.frame_label.pack(side=tk.LEFT, padx=10)

        jump_frame_frame = tk.Frame(nav_frame)
        jump_frame_frame.pack(side=tk.LEFT, padx=10)

        tk.Label(jump_frame_frame, text="Jump to frame:").pack(side=tk.LEFT)

        self.jump_entry = tk.Entry(jump_frame_frame, width=5)
        self.jump_entry.pack(side=tk.LEFT, padx=5)
        self.jump_entry.bind('<Return>', lambda event: self.jump_to_frame())

        tk.Button(jump_frame_frame, text="Go", command=self.jump_to_frame).pack(side=tk.LEFT)

        self.slider = tk.Scale(root, from_=1, to=1, orient=tk.HORIZONTAL, command=self.slider_moved)
        self.slider.pack(fill='x', padx=10, pady=5)
        self.slider.config(state='disabled')

        self.figure, self.ax = plt.subplots()
        self.canvas = FigureCanvasTkAgg(self.figure, master=root)
        self.canvas.get_tk_widget().pack(side='left', fill='both', expand=True)

        self.ax.axis('off')
        self.canvas.draw()

        self.root.bind('<Left>', self.on_left_key)
        self.root.bind('<Right>', self.on_right_key)
        self.root.bind('<Configure>', self.on_resize)

    def load_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("DICOM files", "*.dcm"), ("All files", "*.*")])
        if not filepath:
            return
        try:
            ds = pydicom.dcmread(filepath)
            if not hasattr(ds, 'pixel_array'):
                print("DICOM has no pixel data.")
                return

            self.dicom = ds
            self.pixel_data = ds.pixel_array
            self.num_frames = self.pixel_data.shape[0] if self.pixel_data.ndim == 3 else 1

            self.pixel_spacing = [float(sp) for sp in ds.PixelSpacing] if hasattr(ds, 'PixelSpacing') else [1.0, 1.0]

            self.frame_index = 0

            self.slider.config(to=self.num_frames, state='normal')
            self.slider.set(1)

            filename = os.path.basename(filepath)
            self.filename_label.config(text=f"File: {filename}")

            self.show_frame()

        except Exception as e:
            print(f"Error loading DICOM: {e}")

    def resize_figure(self, frame):
        img_height, img_width = frame.shape
        canvas_height = self.canvas.get_tk_widget().winfo_height()

        desired_width = int(canvas_height * (img_width / img_height))
        dpi = self.figure.dpi
        fig_width_in = desired_width / dpi
        fig_height_in = canvas_height / dpi

        self.figure.set_size_inches(fig_width_in, fig_height_in)

    def show_frame(self):
        if not self.dicom:
            self.slider.config(state='disabled')
            self.ax.clear()
            self.ax.axis('off')
            self.canvas.draw()
            return

        frame = self.pixel_data[self.frame_index] if self.pixel_data.ndim == 3 else self.pixel_data

        self.ax.clear()
        self.ax.axis('off')

        self.resize_figure(frame)
        self.ax.imshow(frame, cmap='gray', aspect='auto')
        self.ax.margins(0)
        self.ax.set_position([0, 0, 1, 1])

        self.frame_label.config(text=f"Frame {self.frame_index + 1} / {self.num_frames}")
        self.slider.set(self.frame_index + 1)
        self.canvas.draw()

    def prev_frame(self):
        if self.dicom and self.frame_index > 0:
            self.frame_index -= 1
            self.show_frame()

    def next_frame(self):
        if self.dicom and self.frame_index < self.num_frames - 1:
            self.frame_index += 1
            self.show_frame()

    def slider_moved(self, val):
        frame_num = int(val) - 1
        if frame_num != self.frame_index:
            self.frame_index = frame_num
            self.show_frame()

    def jump_to_frame(self):
        val = self.jump_entry.get()
        if val.isdigit():
            frame_num = int(val)
            if 1 <= frame_num <= self.num_frames:
                self.frame_index = frame_num - 1
                self.show_frame()

    def on_left_key(self, event):
        self.prev_frame()

    def on_right_key(self, event):
        self.next_frame()

    def on_resize(self, event):
        if self.dicom is None:
            return
        if self.resize_after_id is not None:
            self.root.after_cancel(self.resize_after_id)
        self.resize_after_id = self.root.after(200, self.show_frame)

if __name__ == "__main__":
    root = tk.Tk()
    root.minsize(600, 400)
    app = DICOMViewer(root)
    root.mainloop()