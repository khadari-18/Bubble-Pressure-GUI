import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from collections import deque
import threading
import serial
import serial.tools.list_ports
import numpy as np
import queue
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from datetime import datetime
import time as time_module
from scipy.signal import find_peaks

# Configuration
# SERIAL_PORT = 'dev/cu.usbserial-1140'  # Change to your actual Arduino port
BAUD_RATE = 9600
MAX_DATA_POINTS = 50000
DISPLAY_POINTS = 10000


def prompt_for_excel_path():
    """Prompts the user to choose where to create the Excel file."""
    # Initialize a temporary, hidden root window for the dialog
    temp_root = tk.Tk()
    temp_root.withdraw()

    # Open the native "Save As" file picker
    file_path = filedialog.asksaveasfilename(
        title="Create Excel File Before Launching App",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
    )

    # Clean up the temporary window immediately
    temp_root.destroy()

    # If the user cancels or closes the window, exit the program completely
    if not file_path:
        sys.exit()

    # Attempt to initialize the file with basic headers
    try:
        df = pd.DataFrame(columns=["Timestamp", "Data Input", "Status"])
        df.to_excel(file_path, index=False)
        return file_path
    except Exception as e:
        # Catch errors like permission denied or locked files
        messagebox.showerror("Error", f"Could not create file:\n{e}")
        sys.exit()




class LivePlotApp:
    def __init__(self, root, excel_path=""):
        self.root = root
        self.root.geometry("1080x780") 
        self.root.title("Arduino Live Sensor Plot")
        
        self.ser = None
        self.is_running = True
        self.baseline = 0.0  
        self.corrected_val = 0.0
        self.time_queue = ""

        self.status_queue = queue.Queue()  # thread -> GUI messages
        # self.status_queue.put("Status: Ready")
        self.create_placeholders()      # Call the method to create 2 rows of slots for widget placements 
        self.create_status_bar()        # Call the method to create status bar at the bottom of the window
        self.embed_matplotlib_canvas()  # Call the method to embed the canvas
        self.create_controls_widgets()  # Call the method to create serial port connection controls
        self.create_numerical_widgets()   # Call the method to create zeroing block widgets
        self.create_calibrator_widgets()
        self.create_saving_widgets()    # Call the method to create saving to excel contorls 
        
        self.excel_path = excel_path    # Store the Excel path for later use

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)  # Handling window close events safely
        self.update_liveplot()          # Start the live plot update loop    
        self.update_status()            # Start the status bar update loop
        self.update_numdisplay()        # Start the numerical display update loop


    def create_placeholders(self):
        self.top_frame = tk.Frame(master=self.root, width=1000, height=400)
        self.top_frame.grid(row=0, column=0, sticky="nsew")
        self.top_frame.grid_propagate(False)

        self.bottom_frame = tk.Frame(master=self.root, width=1000, height=380)
        self.bottom_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.bottom_frame.grid_propagate(False)


    def embed_matplotlib_canvas(self):
        self.x_data = list(range(DISPLAY_POINTS))
        self.data_queue = deque(maxlen=MAX_DATA_POINTS)             # Storage buffer is larger. Slice displayed on plot.
        # initialize deque with zeroes to pre-populate the plot
        for _ in range(MAX_DATA_POINTS):
            self.data_queue.append(0)
        self.time_queue = self.data_queue.copy()

        # Setup Matplotlib Plot
        self.fig = Figure(figsize=(6, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_title("Real-Time Sensor Data")
        self.ax.set_ylabel("Sensor Value")
        self.ax.set_xlabel("Time (Samples)")
        self.line, = self.ax.plot(self.x_data, list(self.data_queue)[1:DISPLAY_POINTS+1], 'b-')
        self.peak_dots, = self.ax.plot([], [], 'rv', markersize=8)  # red triangles
        self.ax.grid(True)
        # embed matplotlib canvas into tkinter window
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.top_frame)
        self.canvas.draw()
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.canvas_widget.pack_propagate(False)


    def create_controls_widgets(self):
        # main control block widget setup
        self.main_block = tk.Frame(master=self.top_frame, bg="lightgray", width=200, height=199, bd=2, relief="groove")
        self.main_block.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        self.main_block.grid_propagate(False)
        self.label_controls = tk.Label(master=self.main_block, text="Controls Block", font=("Arial", 14), bg="lightgray")
        self.label_controls.grid(row=0, column=0, padx=5, pady=5)

        # serial port selection in main block widget setup
        self.available_ports = serial.tools.list_ports.comports()
        self.com_ports = [p.device for p in self.available_ports]
        self.selected_com = tk.StringVar(value="-- Select --")
        self.selected_com.trace('w', self.on_port_selected)  # Add trace to detect when selection changes
        self.com_port_menu = tk.OptionMenu(self.main_block, self.selected_com, *(self.com_ports or ["-- No Ports --"]))
        self.com_port_menu.grid(row=1, column=0, padx=5, pady=5)

        self.clear_button = tk.Button(master=self.main_block, text="Clear Data", command=self.clear_data)
        self.clear_button.grid(row=2, column=0, padx=5, pady=5)

        # quit button setup
        self.quit_button = tk.Button(master=self.main_block, text="Quit", command=self.on_close)
        self.quit_button.grid(row=3, column=0, padx=5, pady=5)


    def create_numerical_widgets(self):
        self.zeroing_block = tk.Frame(master=self.bottom_frame, bg="lightgray", width=200, height=250, bd=2, relief="groove")
        self.zeroing_block.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.zeroing_block.grid_propagate(False)
        self.label_zeroing = tk.Label(master=self.zeroing_block, text="Sensor Data", font=("Arial", 14), bg="lightgray")
        self.label_zeroing.grid(row=0, column=0, padx=5, pady=5)

        # saving to excel path input and button setup
        self.press_text = tk.Label(master=self.zeroing_block, text="Pressure Live", font=("Arial", 14), bg="lightgray")
        self.press_text.grid(row=1, column=0, padx=5, pady=5)
        self.press_value = tk.Entry(master=self.zeroing_block, width=10)
        self.press_value.grid(row=2, column=0, padx=5, pady=5)
        self.press_unit = tk.Label(master=self.zeroing_block, text="Pa", font=("Arial", 14), bg="lightgray")
        self.press_unit.grid(row=2, column=1, padx=5, pady=5)

        self.zero_text = tk.Label(master=self.zeroing_block, text="Current Baseline", font=("Arial", 14), bg="lightgray")
        self.zero_text.grid(row=3, column=0, padx=5, pady=5)
        self.zero_button = tk.Button(master=self.zeroing_block, text="Zero Pressure", command=self.set_baseline)
        self.zero_button.grid(row=4, column=0, padx=5, pady=5)
        self.zero_value = tk.Entry(master=self.zeroing_block, width=10)
        self.zero_value.grid(row=5, column=0, padx=5, pady=5)
        self.press_unit = tk.Label(master=self.zeroing_block, text="Pa", font=("Arial", 14), bg="lightgray")
        self.press_unit.grid(row=5, column=1, padx=5, pady=5)


    def create_calibrator_widgets(self):
        # saving to excel block widget setup
        self.calib_block = tk.Frame(master=self.bottom_frame, bg="lightgray", width=250, height=250, bd=2, relief="groove")
        self.calib_block.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        self.calib_block.grid_propagate(False)
        self.label_calib = tk.Label(master=self.calib_block, text="Calibration Block", font=("Arial", 14), bg="lightgray")
        self.label_calib.grid(row=0, column=0, padx=5, pady=5)

        # saving to excel path input and button setup
        self.calib_text = tk.Label(master=self.calib_block, text="Input Calibrator Tension", font=("Arial", 14), bg="lightgray")
        self.calib_text.grid(row=1, column=0, padx=5, pady=5)
        self.tension_value = tk.Entry(master=self.calib_block, width=10)
        self.tension_value.grid(row=2, column=0, padx=5, pady=5)
        self.tension_value.insert(0,"72.1")
        self.tension_unit = tk.Label(master=self.calib_block, text="m N/m", font=("Arial", 14), bg="lightgray")
        self.tension_unit.grid(row=2, column=1, padx=5, pady=5)
        self.calib_button = tk.Button(master=self.calib_block, text="Calculate Radius", command=self.calc_radius)
        self.calib_button.grid(row=3, column=0, padx=5, pady=5)
        self.radius_out = tk.Entry(master=self.calib_block, width=10)
        self.radius_out.grid(row=4, column=0, padx=5, pady=5)
        self.radius_text = tk.Label(master=self.calib_block, text="mm", font=("Arial", 14), bg="lightgray")
        self.radius_text.grid(row=4, column=1, padx=5, pady=5)
        self.prom_text = tk.Label(master=self.calib_block, text="Set Prominance", font=("Arial", 14), bg="lightgray") 
        self.prom_text.grid(row=5, column=0, padx=5, pady=5)
        self.prominance_value = tk.Entry(master=self.calib_block, width=10)
        self.prominance_value.grid(row=5, column=1, padx=5, pady=5)
        self.prominance_value.insert(0,"10")            # initially set prominance to 10, user can change it to any value.


    def create_saving_widgets(self):
        # saving to excel block widget setup
        self.excel_block = tk.Frame(master=self.bottom_frame, bg="lightgray", width=200, height=250, bd=2, relief="groove")
        self.excel_block.grid(row=0, column=2, sticky="ew", padx=5, pady=5)
        self.excel_block.grid_propagate(False)
        self.label_savings = tk.Label(master=self.excel_block, text="Excel Block", font=("Arial", 14), bg="lightgray")
        self.label_savings.grid(row=0, column=0, padx=5, pady=5)

        # # saving to excel path input and button setup
        self.excel_sheet = tk.Entry(master=self.excel_block, width=15)
        self.excel_sheet.grid(row=1, column=0, padx=5, pady=5)
        self.excel_sheet.insert(0,"-- Sheet Name --")
        self.excel_column = tk.Entry(master=self.excel_block, width=15)
        self.excel_column.grid(row=2, column=0, padx=5, pady=5)
        self.excel_column.insert(0,"-- Column Name --")
        self.save_button = tk.Button(master=self.excel_block, text="Save Peaks", command=self.save_to_excel)
        self.save_button.grid(row=3, column=0, padx=5, pady=5)


    def create_status_bar(self):
        # status bar widget setup
        self.status_bar = tk.Frame(master=self.bottom_frame, bg="lightgray", width=610, height=30)
        self.status_bar.grid(row=1, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        self.status_bar.grid_propagate(False)
        
        # status bar label setup
        self.label_status = tk.Label(master=self.status_bar, text="Status: Ready", font=("Arial", 12), bg="lightgray")
        self.label_status.grid(row=0, column=0, padx=5, pady=5)
        

    def on_port_selected(self, *args):
        """Called when a port is selected from the dropdown."""
        if self.selected_com.get() != "-- Select --":
            self.connect_serial()


    def connect_serial(self):
        """Attempts to connect to the selected serial port and starts the data reading thread."""
        if self.ser and self.ser.is_open:
            try:
                self.ser.close()
            except Exception as e:
                self.status_queue.put(f"Error closing serial port: {e}")
                self.ser = None

        try:
            self.ser = serial.Serial(self.selected_com.get(), BAUD_RATE, timeout=0.1)
            self.serial_thread = threading.Thread(target=self.read_serial_data, daemon=True)  # Threading is better for speed? 
            self.serial_thread.start()
        except Exception as e:
            self.status_queue.put(f"Error opening serial port {self.selected_com.get()}: {e}")  


    def read_serial_data(self):
        """Runs in a separate background thread to collect data."""
        while self.is_running:
            if not self.ser or not self.ser.is_open:
                time_module.sleep(0.05)  # Avoid busy waiting if serial is not connected
                continue

            try:
                if self.ser.in_waiting == 0:
                    time_module.sleep(0.05)  # Avoid busy waiting if serial is not connected
                    continue

                val = self.ser.readline() # .decode('utf-8').strip() not needed as code uploaded to arduino already does that to serial data before sending here. 
                # time_module.sleep(0.01)  # Small delay to prevent CPU overuse
                self.corrected_val = float(val) - self.baseline
                now = datetime.now().time().strftime("%H:%M:%S.%f")[:-3] 
                self.data_queue.append(self.corrected_val)
                self.time_queue.append(now)

            except (ValueError, serial.SerialException):
                # Handle temporary bad strings or disconnected hardware safely
                continue


    def update_numdisplay(self):
        """Updates the numeric display of the latest sensor value."""
        self.press_value.config(state="normal")
        self.press_value.delete(0, tk.END)
        self.press_value.insert(0, str(self.data_queue[-1]))
        self.press_value.config(state="readonly") # Prevents user typing over it
        self.root.after(100, self.update_numdisplay)


    def update_liveplot(self):
        """Runs in the main Tkinter thread to safely redraw the UI."""
        if not self.is_running:
            return
        
        display_data = np.asarray(list(self.data_queue)[-DISPLAY_POINTS:], dtype=float)
        self.line.set_data(self.x_data, display_data)        # Update line data 

        # Refresh peak markers every 10th frame (~5 times/sec)
        self._frame_count = getattr(self, "_frame_count", 0) + 1
        if self._frame_count % 10 == 0:
            indices, _ = find_peaks(display_data, prominence=float(self.prominance_value.get()))
            if indices.size > 0:
                self.peak_dots.set_data(indices, display_data[indices]) # plotting x,y of peaks
            else:
                self.peak_dots.set_data([], [])
                
        self.ax.relim(); self.ax.autoscale_view()       # Rescale axes to fit new data
        self.canvas.draw_idle()                         # Efficiently wipe out canvas for new data
        self.root.after(20, self.update_liveplot)       # Schedule next frame update in 20 milliseconds (~50 FPS)
        

    def update_status(self):
        """Runs in the main Tkinter thread to safely redraw the UI."""
        if not self.is_running: # protection against callbacks while gui not running?
            return
        try:
            # Show the most recent queued message, if any
            msg = self.status_queue.get_nowait()
            self.label_status.config(text=msg)
            self._status_hold = 25  # hold custom message for ~50 cycles (5 s at 100 ms)
        except queue.Empty:                             # when system idle, either show connected or not by default 
            if getattr(self, "_status_hold", 0) > 0:
                self._status_hold -= 1  # keep the custom message visible
            elif self.ser and self.ser.is_open:
                self.label_status.config(text="Connected. System reading.")
            else:
                self.label_status.config(text="No serial connected.")
        self.root.after(100, self.update_status)         # Schedule next frame update in 100 milliseconds


    def save_to_excel(self):
        """Saves read serial data to the specified excel path."""
        peaks, times = self.get_peaks()
        if len(peaks) == 0:
            self.status_queue.put("No peaks could be detected. Could not save data.")
            return
        
        sheet = self.excel_sheet.get()
        try:
            data_df = pd.read_excel(self.excel_path, sheet_name=sheet, engine="openpyxl")
        except (FileNotFoundError, ValueError):
            self.status_queue.put("Saving first set of data.")
            data_df = pd.DataFrame()
        _, col = data_df.shape
        
        sample_num = col // 3 + 1
        new_df = pd.DataFrame({
            f"Sample {sample_num}": [None]*len(peaks),
            f"Timestamp {sample_num}": times,
            f"Data Input {sample_num}": peaks
        })
        new_df.iloc[0, 1] = self.excel_column.get()

        data_df = pd.concat([data_df, new_df], axis=1)
        try:
            with pd.ExcelWriter(self.excel_path, mode="a", engine="openpyxl",
                if_sheet_exists="replace") as writer:
                data_df.to_excel(writer, index=False, sheet_name=sheet)
        except PermissionError:
            self.status_queue.put("Could not write — is the Excel file open in another program?")

    def set_baseline(self):
        """Sets the baseline to the last value in the data queue."""
        self.baseline = self.baseline + self.corrected_val

        self.zero_value.config(state="normal")
        self.zero_value.delete(0, tk.END)
        self.zero_value.insert(0, f"{self.baseline:.2f}")
        self.zero_value.config(state="readonly") # Prevents user typing over it


    def clear_data(self):
        """Clears the data queue and resets the plot."""
        self.data_queue.clear()
        self.data_queue.extend([0] * MAX_DATA_POINTS)
        self.time_queue = self.data_queue.copy()    # putting zeros in time_queue 
        self.status_queue.put("Data cleared from queue.")


    def get_peaks(self):
        """Detects peaks in the data queue and returns their indices and values."""
        peaks = []
        if all(x == 0 for x in self.data_queue):    # data_eque is filled with zeros at launching of the program.    
            return [],[]
        
        data = np.asarray(self.data_queue, dtype=float)
        indices, _ = find_peaks(data, prominence=float(self.prominance_value.get()))

        if indices.size == 0:
            return [],[]

        peaks = data[indices]
        times = np.asarray(self.time_queue)[indices]

        if len(peaks) < 5: 
            self.status_queue.put("Warning: Less than 5 peaks detected.")

        return peaks, times


    def calc_radius(self):
        """Calculates the radius based on the detected peaks."""
        peaks, _ = self.get_peaks()
        if len(peaks) == 0:
            return
        
        avg_p = sum(peaks)/len(peaks)
        if avg_p == 0.0 or not avg_p:
            self.status_queue.put("Null peaks read, Couldn't calculate radius.")
            return

        radius = 2 * float(self.tension_value.get())/avg_p

        # Enable editing, clear old value, insert new value, disable editing
        self.radius_out.config(state="normal")
        self.radius_out.delete(0, tk.END)
        self.radius_out.insert(0, str(radius))
        self.radius_out.config(state="readonly") # Prevents user typing over it


    def on_close(self):
        """Clean closure of threads and serial ports when exiting."""
        self.is_running = False
        if self.ser and self.ser.is_open:
            self.ser.close()
        self.root.destroy()


if __name__ == "__main__":
    path = prompt_for_excel_path()
    root = tk.Tk()
    app = LivePlotApp(root, excel_path=path)
    root.mainloop()

