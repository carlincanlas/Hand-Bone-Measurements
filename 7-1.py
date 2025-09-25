import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pydicom
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from openpyxl import Workbook
import ctypes

### PREVIOUS MEASUREMENT METHOD (not accurate but clean working code) ###
# Windows compatibility
# DIPI awareness for high DPI displays
# Slope remains on screen and is less opaque
# Copy & Export measurements
# Right column button panel

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

class DICOMViewer:
    def __init__(self, root):
        # Main root window
        self.root = root
        self.root.title("Hand DICOM Viewer with Measurements")
        self.frame_index = 0
        self.num_frames = 1
        self.dicom = None
        self.pixel_data = None
        self.pixel_spacing = [1.0, 1.0]
        self.resize_after_id = None

        # Measurement tools
        self.measurements = {}
        self.points = []
        self.selected_point = None
        self.dragging = False
        self.bone_lines = {}
        self.bone_slope = {}
        self.h_next_joint_x = None
        self.measure_step = None

        # Zoom in & pan 
        self.zoom_level = 0         # Zoom in from 0% to 100%
        self.pan_offset = [0, 0]    # [x_offset, y_offset]
        self.is_panning = False
        self.last_pan_xy = None

        # Copy measurements to range of frames
        self.control_frame = tk.Frame(self.root)
        self.control_frame.pack(side=tk.RIGHT, fill=tk.Y)

        # 2 columns:
        # DICOM canvas on the left
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill='both', expand=True)
        self.main_frame.rowconfigure(0, weight=1)
        # Matplotlib figure and canvas
        self.figure, self.ax = plt.subplots()
        self.ax.axis('off')
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.main_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky='nsew')
        # Controls frame on the right side
        self.control_frame = tk.Frame(self.main_frame, width=250)
        self.control_frame.grid(row=0, column=1, sticky='ns')
        self.control_frame.grid_propagate(False)  # Fix width so it doesn't shrink

        # Configure grid weights so canvas expands and controls stay fixed width
        self.main_frame.columnconfigure(0, weight=1)  # Canvas column expands horizontally
        self.main_frame.columnconfigure(1, weight=0)  # Controls fixed width
        self.main_frame.rowconfigure(0, weight=1)     # Row expands vertically

        # Controls - grouped in frames for neat layout
        file_frame = tk.Frame(self.control_frame)
        file_frame.pack(pady=5, fill='x')

        self.filename_label = tk.Label(file_frame, text="No file loaded", anchor='w')
        self.filename_label.pack(side=tk.LEFT, padx=(0, 10), fill='x', expand=True)

        tk.Button(file_frame, text="Load DICOM File", command=self.load_file).pack(side=tk.LEFT)

        measure_frame = tk.Frame(self.control_frame)
        measure_frame.pack(pady=5, fill='x')

        tk.Button(measure_frame, text="Measure", command=self.start_measurement_workflow).pack(side=tk.LEFT, pady=2, padx=2, fill='x', expand=True)
        tk.Button(measure_frame, text="Clear Measurements", command=self.clear_measurements).pack(side=tk.LEFT, pady=2, padx=2, fill='x', expand=True)
        tk.Button(measure_frame, text="Export Measurements", command=self.export_measurements).pack(side=tk.LEFT, pady=2, padx=2, fill='x', expand=True)

        copy_frame = tk.Frame(self.control_frame)
        copy_frame.pack(pady=5, fill='x')
        tk.Label(copy_frame, text="Copy to frames (eg. 1-5):").pack(side=tk.LEFT)
        self.range_entry = tk.Entry(copy_frame, width=10)
        self.range_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(copy_frame, text="Copy", command=self.copy_measurements_to_range).pack(side=tk.LEFT)

        self.label = tk.Label(self.control_frame, text="", anchor='w', justify='left')
        self.label.pack(pady=5, fill='x')

        self.slider = tk.Scale(self.control_frame, from_=1, to=1, orient=tk.HORIZONTAL, command=self.slider_moved)
        self.slider.pack(fill='x', padx=10, pady=5)
        self.slider.config(state='disabled')

        nav_frame = tk.Frame(self.control_frame)
        nav_frame.pack(pady=5, fill='x')

        tk.Button(nav_frame, text="Previous", command=self.prev_frame).pack(side=tk.LEFT, padx=2, fill='x', expand=True)
        tk.Button(nav_frame, text="Next", command=self.next_frame).pack(side=tk.LEFT, padx=2, fill='x', expand=True)

        self.frame_label = tk.Label(nav_frame, text="Frame 1 / 1", width=15)
        self.frame_label.pack(side=tk.LEFT, padx=10)

        jump_frame_frame = tk.Frame(self.control_frame)
        jump_frame_frame.pack(pady=5, fill='x')

        tk.Label(jump_frame_frame, text="Jump to frame:").pack(side=tk.LEFT)

        self.jump_entry = tk.Entry(jump_frame_frame, width=5)
        self.jump_entry.pack(side=tk.LEFT, padx=5)
        self.jump_entry.bind('<Return>', lambda event: self.jump_to_frame())

        tk.Button(jump_frame_frame, text="Go", command=self.jump_to_frame).pack(side=tk.LEFT)

        zoom_frame = tk.Frame(self.control_frame)
        zoom_frame.pack(pady=5, fill='x')
        tk.Label(zoom_frame, text="Zoom:").pack(side=tk.LEFT)
        self.zoom_slider = tk.Scale(zoom_frame, from_=0, to=100, orient=tk.HORIZONTAL, command=self.on_zoom_change)
        self.zoom_slider.pack(side=tk.LEFT, padx=5, fill='x', expand=False)
        tk.Button(zoom_frame, text="Reset View", command=self.reset_zoom).pack(side=tk.LEFT)

        # Bind events
        self.canvas.mpl_connect("button_press_event", self.on_mouse_press)
        self.canvas.mpl_connect("motion_notify_event", self.on_mouse_move)
        self.canvas.mpl_connect("button_release_event", self.on_mouse_release)
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
            self.measurements.clear()
            self.points.clear()
            self.selected_point = None
            self.dragging = False
            self.pan_offset = [0, 0]
            self.zoom_level = 0
            self.zoom_slider.set(0)

            self.slider.config(to=self.num_frames, state='normal')
            self.slider.set(1)

            filename = os.path.basename(filepath)
            self.filename_label.config(text=f"File: {filename}")

            self.show_frame()

        except Exception as e:
            print(f"Error loading DICOM: {e}")


    def show_frame(self):
        if not self.dicom:
            self.slider.config(state='disabled')
            self.ax.clear()
            self.ax.axis('off')
            self.canvas.draw_idle()
            return

        frame = self.pixel_data[self.frame_index] if self.pixel_data.ndim == 3 else self.pixel_data
        self.ax.clear()
        self.ax.axis('off')

        self.ax.imshow(frame, cmap='gray', aspect='equal', extent=[0, frame.shape[1], frame.shape[0], 0])
        self.figure.subplots_adjust(left=0, right=1, top=1, bottom=0)


        # Set limits to show entire image
        self.ax.set_xlim(0, frame.shape[1])
        self.ax.set_ylim(frame.shape[0], 0)  # flip y axis for correct orientation

        # Apply zoom and pan:
        img_height, img_width = frame.shape
        zoom_fraction = self.zoom_level / 100.0
        zoom_width = img_width * (1 - zoom_fraction)
        zoom_height = img_height * (1 - zoom_fraction)

        center_x = img_width / 2 + self.pan_offset[0]
        center_y = img_height / 2 + self.pan_offset[1]

        x0 = max(0, center_x - zoom_width / 2)
        x1 = min(img_width, center_x + zoom_width / 2)
        y0 = max(0, center_y - zoom_height / 2)
        y1 = min(img_height, center_y + zoom_height / 2)

        self.ax.set_xlim(x0, x1)
        self.ax.set_ylim(y1, y0)  # flip y axis

        # Draw measurements overlays
        dot = 'o'
        dotsize = 0.5
        linesize = 0.8

        frame_measures = self.measurements.get(self.frame_index, {})
        for key, color in [('h', 'red'), ('H', 'yellow')]:
            if key in frame_measures:
                p1, p2 = frame_measures[key]
                self.ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color=color, linewidth=linesize)
                self.ax.plot(p1[0], p1[1], marker=dot, markersize=dotsize, color=color)
                self.ax.plot(p2[0], p2[1], marker=dot, markersize=dotsize, color=color)

        # Draw the bone line (cyan dashed) if it was confirmed for this frame
        if self.frame_index in self.bone_lines:
            p1, p2 = self.bone_lines[self.frame_index]
            self.ax.plot(
                [p1[0], p2[0]], [p1[1], p2[1]],
                linestyle='dashed',
                linewidth=linesize,
                color='cyan',
                alpha=0.4)  # adjust this value between 0 (invisible) and 1 (fully opaque)
            self.ax.plot(p1[0], p1[1], marker=dot, markersize=dotsize, color='cyan')
            self.ax.plot(p2[0], p2[1], marker=dot, markersize=dotsize, color='cyan')

        if len(self.points) == 1:
            self.ax.plot(self.points[0][0], self.points[0][1], marker=dot, markersize=dotsize, color='cyan')
        elif len(self.points) == 2:
            p1, p2 = self.points
            self.ax.plot([p1[0], p2[0]], [p1[1], p2[1]], linestyle='dashed', linewidth=linesize, color='cyan')
            self.ax.plot(p1[0], p1[1], marker=dot, markersize=dotsize, color='cyan')
            self.ax.plot(p2[0], p2[1], marker=dot, markersize=dotsize, color='cyan')

        self.frame_label.config(text=f"Frame {self.frame_index + 1} / {self.num_frames}")
        self.slider.set(self.frame_index + 1)

        self.canvas.draw_idle()
        self.update_measurement_label()


    def ask_bone_line_confirmation(self):
        confirm = messagebox.askyesno("Confirm", "Confirm bone line?")
        if not confirm:
            self.points.clear()
            self.measure_step = 'bone_start'
            self.label.config(text="Step 1: Click left side of bone line")
            self.show_frame()
            return

        dx = self.points[1][0] - self.points[0][0]
        dy = self.points[1][1] - self.points[0][1]
        slope = dy / dx if dx != 0 else float('inf')
        self.bone_lines[self.frame_index] = tuple(self.points)
        self.bone_slope[self.frame_index] = slope
        self.label.config(text="Step 3: Click start of h measurement (next joint)")
        self.points.clear()
        self.measure_step = 'h_start'
        self.show_frame()


    def start_measurement_workflow(self):
        self.measure_step = 'bone_start'
        self.points.clear()
        self.label.config(text="Step 1: Click left side of bone line")


    def on_mouse_press(self, event):
        if not self.dicom or event.inaxes != self.ax:
            return
        x, y = event.xdata, event.ydata

        # Check for Shift + drag to pan
        if event.key == 'shift':
            self.is_panning = True
            self.last_pan_xy = (event.x, event.y)
            return


        if self.measure_step == 'bone_start':
            self.points = [(x, y)]
            self.label.config(text="Step 2: Click right side of bone line")
            self.measure_step = 'bone_end'
            return

        elif self.measure_step == 'bone_end':
            self.points.append((x, y))
            self.show_frame()  # Draw second point and preview line
            # Delay the confirmation so the GUI can update
            self.root.after(10, self.ask_bone_line_confirmation)
            return

        elif self.measure_step == 'h_start':
            self.points = [(x, y)]
            self.label.config(text="Step 4: Click next joint")
            self.measure_step = 'h_end'
            return

        elif self.measure_step == 'h_end':
            self.points.append((x, y))
            x1 = self.points[0][0]
            x2 = self.points[1][0]
            slope = self.bone_slope.get(self.frame_index, 0)
            y0 = self.points[0][1]
            h_line = ((x1, y0), (x2, y0 + (x2 - x1) * slope))
            self.h_next_joint_x = x2
            if self.frame_index not in self.measurements:
                self.measurements[self.frame_index] = {}
            self.measurements[self.frame_index]['h'] = h_line
            self.label.config(text="Step 5: Click base of epiphysis")
            self.measure_step = 'H_base'
            self.points.clear()
            self.show_frame()
            return

        elif self.measure_step == 'H_base':
            x1 = x
            y1 = y
            x2 = self.h_next_joint_x
            slope = self.bone_slope.get(self.frame_index, 0)
            dy = (x2 - x1) * slope
            offset = 20
            H_line = ((x1, y1 + offset), (x2, y1 + dy + offset))
            self.measurements[self.frame_index]['H'] = H_line
            self.measure_step = None
            self.show_frame()
            return

        frame_measures = self.measurements.get(self.frame_index, {})
        threshold_line = 15   # bigger threshold for clicking/dragging the line
        threshold_endpoints = 5  # smaller threshold for endpoints

        def dist_sq(a, b):
            return (a[0] - b[0])**2 + (a[1] - b[1])**2

        for key, (p1, p2) in frame_measures.items():
            if dist_sq((x, y), p1) < threshold_endpoints**2:
                self.selected_point = (key, 'p1')
                self.dragging = True
                return
            if dist_sq((x, y), p2) < threshold_endpoints**2:
                self.selected_point = (key, 'p2')
                self.dragging = True
                return
            if self.point_near_line((x, y), p1, p2, threshold_line):
                self.selected_point = (key, 'line')
                self.dragging = True
                self.drag_offset = (x, y)
                return



    def on_mouse_move(self, event):

        if self.is_panning and self.last_pan_xy:
            dx = event.x - self.last_pan_xy[0]
            dy = event.y - self.last_pan_xy[1]
            self.last_pan_xy = (event.x, event.y)

            img_height, img_width = self.pixel_data[0].shape if self.pixel_data.ndim == 3 else self.pixel_data.shape
            zoom_fraction = self.zoom_level / 100.0
            visible_width = img_width * (1 - zoom_fraction)
            visible_height = img_height * (1 - zoom_fraction)

            # Convert screen pixels to data units
            bbox = self.ax.get_window_extent().transformed(self.figure.dpi_scale_trans.inverted())
            ax_width, ax_height = bbox.width * self.figure.dpi, bbox.height * self.figure.dpi

            if ax_width == 0 or ax_height == 0:
                return

            data_dx = dx * visible_width / ax_width
            data_dy = dy * visible_height / ax_height

            self.pan_offset[0] -= data_dx
            self.pan_offset[1] += data_dy  # Inverted y-axis

            self.show_frame()
            return

        if not self.dicom or not self.dragging or event.inaxes != self.ax or not self.selected_point:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        key, part = self.selected_point
        p1, p2 = self.measurements[self.frame_index][key]
        slope = self.bone_slope.get(self.frame_index, 0)

        if part == 'p1':
            dy = (p2[0] - x) * slope
            new_p1 = (x, p2[1] - dy)
            self.measurements[self.frame_index][key] = (new_p1, p2)
        elif part == 'p2':
            dx = x - p2[0]
            dy = y - p2[1]
            new_p2 = (p2[0] + dx, p2[1] + dy)
            new_p2_x = new_p2[0]
            new_p2_y = p1[1] + (new_p2_x - p1[0]) * slope
            self.measurements[self.frame_index][key] = (p1, (new_p2_x, new_p2_y))

            # Sync the x2 of the other measurement if both exist
            other_key = 'H' if key == 'h' else 'h'
            if other_key in self.measurements[self.frame_index]:
                op1, op2 = self.measurements[self.frame_index][other_key]
                ody = (new_p2_x - op1[0]) * slope
                synced_op2 = (new_p2_x, op1[1] + ody)
                self.measurements[self.frame_index][other_key] = (op1, synced_op2)

        elif part == 'line':
            dy = y - self.drag_offset[1]
            new_p1 = (p1[0], p1[1] + dy)
            new_p2 = (p2[0], p2[1] + dy)
            self.measurements[self.frame_index][key] = (new_p1, new_p2)
            self.drag_offset = (x, y)

        self.show_frame()


    def on_mouse_release(self, event):
        if self.is_panning:
            self.is_panning = False
            self.last_pan_xy = None
            return

        self.dragging = False
        self.selected_point = None
        self.show_frame()

    def point_near_line(self, pt, line_start, line_end, threshold):
        x, y = pt
        x1, y1 = line_start
        x2, y2 = line_end
        if (x1 == x2) and (y1 == y2):
            return False
        vx, vy = x2 - x1, y2 - y1
        wx, wy = x - x1, y - y1
        c1 = vx * wx + vy * wy
        if c1 <= 0:
            return (x - x1)**2 + (y - y1)**2 <= threshold**2
        c2 = vx * vx + vy * vy
        if c2 <= c1:
            return (x - x2)**2 + (y - y2)**2 <= threshold**2
        b = c1 / c2
        pbx = x1 + b * vx
        pby = y1 + b * vy
        return (x - pbx)**2 + (y - pby)**2 <= threshold**2


    def calculate_distance(self, p1, p2):
        dx = (p2[0] - p1[0]) * self.pixel_spacing[0]
        dy = (p2[1] - p1[1]) * self.pixel_spacing[1]
        return (dx ** 2 + dy ** 2) ** 0.5

    def update_measurement_label(self):
        frame_measures = self.measurements.get(self.frame_index, {})
        h_dist = H_dist = None
        if 'h' in frame_measures:
            h_dist = self.calculate_distance(*frame_measures['h'])
        if 'H' in frame_measures:
            H_dist = self.calculate_distance(*frame_measures['H'])

        if h_dist and H_dist and self.measure_step is None:
            self.label.config(text=f"Measurements complete. Drag to adjust.  h: {h_dist:.2f} mm  |  H: {H_dist:.2f} mm")
        elif self.measure_step is not None:
            self.label.config(text=self.label.cget("text"))
        else:
            self.label.config(text="")


    def clear_measurements(self):
        confirm = messagebox.askyesno("Confirm", "Clear all measurements for this frame?")
        if not confirm:
            return

        if self.frame_index in self.measurements:
            del self.measurements[self.frame_index]
        if self.frame_index in self.bone_lines:
            del self.bone_lines[self.frame_index]
        if self.frame_index in self.bone_slope:
            del self.bone_slope[self.frame_index]

        self.points.clear()
        self.selected_point = None
        self.dragging = False
        self.measure_step = None
        self.h_next_joint_x = None
        self.label.config(text="")
        self.show_frame()


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


    def export_measurements(self):
        if not self.measurements:
            messagebox.showinfo("No Data", "There are no measurements to export.")
            return

        # Extract base filename without extension (if available)
        if self.dicom and hasattr(self.dicom, 'filename'):
            base_filename = os.path.splitext(os.path.basename(self.dicom.filename))[0]
        else:
            base_filename = "measurements"

        suggested_name = base_filename + ".xlsx"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=suggested_name,
            title="Save Measurements As"
        )

        if not save_path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Measurements"
        ws.append(["Filename", "Frame", "h (mm)", "H (mm)", "Bone Slope"])

        filename = os.path.basename(self.dicom.filename) if self.dicom and hasattr(self.dicom, 'filename') else "Unknown"

        for frame_num in sorted(self.measurements.keys()):
            frame_data = self.measurements[frame_num]
            slope = self.bone_slope.get(frame_num, None)

            h_dist = None
            H_dist = None

            if 'h' in frame_data:
                h_dist = round(self.calculate_distance(*frame_data['h']), 2)
            if 'H' in frame_data:
                H_dist = round(self.calculate_distance(*frame_data['H']), 2)

            ws.append([
                filename,
                frame_num + 1,  # Convert 0-indexed to 1-based
                h_dist if h_dist is not None else "",
                H_dist if H_dist is not None else "",
                slope if slope is not None else ""
            ])

        try:
            wb.save(save_path)
            messagebox.showinfo("Success", f"Measurements exported to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file:\n{e}")

    def copy_measurements_to_range(self):
        if self.frame_index not in self.measurements:
            messagebox.showinfo("No Data", "No measurements found on current frame.")
            return

        range_str = self.range_entry.get().strip()
        if not range_str:
            return

        try:
            start_str, end_str = range_str.split('-')
            start = int(start_str) - 1  # Convert to 0-based index
            end = int(end_str) - 1
        except ValueError:
            messagebox.showerror("Invalid Input", "Enter range as start-end (e.g. 5-12)")
            return

        if start > end or start < 0 or end >= self.num_frames:
            messagebox.showerror("Out of Range", f"Enter a valid range between 1 and {self.num_frames}")
            return

        source_frame = self.frame_index
        source_measures = {}
        if source_frame in self.measurements:
            if 'h' in self.measurements[source_frame]:
                source_measures['h'] = self.measurements[source_frame]['h'][:]
            if 'H' in self.measurements[source_frame]:
                source_measures['H'] = self.measurements[source_frame]['H'][:]

        source_bone = self.bone_lines.get(source_frame)
        source_slope = self.bone_slope.get(source_frame)

        for i in range(start, end + 1):
            self.measurements[i] = {}
            if 'h' in source_measures:
                self.measurements[i]['h'] = source_measures['h'][:]
            if 'H' in source_measures:
                self.measurements[i]['H'] = source_measures['H'][:]
            if source_bone:
                self.bone_lines[i] = source_bone[:]
            if source_slope is not None:
                self.bone_slope[i] = source_slope

        messagebox.showinfo("Success", f"Measurements copied to frames {start+1} to {end+1}.")

    def on_resize(self, event):
        if self.dicom is None:
            return
        if self.resize_after_id is not None:
            self.root.after_cancel(self.resize_after_id)
        self.resize_after_id = self.root.after(200, self.show_frame)


    def on_zoom_change(self, val):
        self.zoom_level = int(val)
        self.show_frame()


    def reset_zoom(self):
        self.zoom_level = 0
        self.pan_offset = [0, 0]
        self.zoom_slider.set(0)
        self.show_frame()


if __name__ == "__main__":
    root = tk.Tk()
    viewer = DICOMViewer(root)
    root.geometry("1200x700")
    root.mainloop()
