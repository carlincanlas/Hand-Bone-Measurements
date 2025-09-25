import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import pydicom

def open_dicom():
    filename = filedialog.askopenfilename(
        title="Select DICOM file",
        filetypes=[("DICOM files", "*.dcm"), ("All files", "*.*")]
    )
    if not filename:
        return

    try:
        ds = pydicom.dcmread(filename)

        # Gather metadata
        info = []
        info.append(f"File: {filename}")
        info.append(f"Modality: {getattr(ds, 'Modality', 'N/A')}")
        info.append(f"Rows: {getattr(ds, 'Rows', 'N/A')}")
        info.append(f"Columns: {getattr(ds, 'Columns', 'N/A')}")
        info.append(f"Number of Frames: {getattr(ds, 'NumberOfFrames', '1')}")

        if "PixelSpacing" in ds:
            row_spacing, col_spacing = ds.PixelSpacing
            info.append(f"Pixel Spacing: Row={row_spacing} mm, Col={col_spacing} mm")
        else:
            info.append("Pixel Spacing: Not found")

        if "SliceThickness" in ds:
            info.append(f"Slice Thickness: {ds.SliceThickness} mm")

        if "SpacingBetweenSlices" in ds:
            info.append(f"Spacing Between Slices: {ds.SpacingBetweenSlices} mm")

        # Show in text box
        text_box.delete("1.0", tk.END)
        text_box.insert(tk.END, "\n".join(info))

    except Exception as e:
        messagebox.showerror("Error", f"Failed to read DICOM: {e}")

# GUI setup
root = tk.Tk()
root.title("DICOM Metadata Viewer")
root.geometry("600x400")

btn = tk.Button(root, text="Open DICOM File", command=open_dicom)
btn.pack(pady=10)

text_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=20)
text_box.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

root.mainloop()
