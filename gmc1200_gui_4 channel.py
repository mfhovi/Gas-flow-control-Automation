import threading
import time
import math
import csv

# Tkinter import is optional so that the module can still be imported
# in environments where tkinter is not available (e.g. headless/test sandbox).
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
except ModuleNotFoundError:
    tk = None
    ttk = None
    messagebox = None
    filedialog = None

try:
    import serial
    from serial.tools import list_ports
except ImportError:
    serial = None
    list_ports = None

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class GMC1200Controller:
    def __init__(self):
        self.ser = None
        self.lock = threading.Lock()
        self.connected = False

    def connect(self, port, baudrate=9600, timeout=0.2):
        if serial is None:
            raise RuntimeError("pyserial is not installed.")
        if self.ser and self.ser.is_open:
            self.ser.close()
        self.ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=timeout,
        )
        self.connected = True

    def disconnect(self):
        if self.ser and self.ser.is_open:
            self.ser.close()
        self.connected = False

    def send_cmd(self, cmd, expect_response=True):
        if not self.connected or not self.ser:
            raise RuntimeError("Not connected to GMC1200.")
        if not cmd.startswith("#"):
            cmd = "#" + cmd
        msg = (cmd + "\r").encode("ascii")
        with self.lock:
            self.ser.reset_input_buffer()
            self.ser.write(msg)
            if not expect_response:
                return ""
            time.sleep(0.05)
            resp = self.ser.read(100)
        return resp.decode(errors="ignore").strip()

    def set_setpoint(self, ch, value):
        cmd = f"SS{ch} {value}"
        return self.send_cmd(cmd)

    def set_flow_on_off(self, ch, on):
        val = 1 if on else 0
        cmd = f"SF{ch} {val}"
        return self.send_cmd(cmd)

    def read_flow(self, ch):
        cmd = f"RF{ch}"
        resp = self.send_cmd(cmd)
        return resp

    def all_off(self, max_ch=8):
        for ch in range(1, max_ch + 1):
            try:
                self.set_flow_on_off(ch, False)
            except Exception:
                pass


if tk is not None:

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("GMC1200 Flow Controller")
            self.geometry("1100x800")
            self.controller = GMC1200Controller()

            self.poll_interval_ms = 1000  # flow polling interval

            # Current setpoints (for plotting)
            self.current_setpoint_A = 0.0
            self.current_setpoint_B = 0.0
            self.current_setpoint_C = 0.0
            self.current_setpoint_D = 0.0

            # Logical channel mapping: A, B, C, D -> physical channel numbers
            self.chA_var = tk.StringVar(value="1")
            self.chB_var = tk.StringVar(value="2")
            self.chC_var = tk.StringVar(value="3")
            self.chD_var = tk.StringVar(value="4")

            # Manual setpoints
            self.setpointA_var = tk.DoubleVar(value=0.0)
            self.setpointB_var = tk.DoubleVar(value=0.0)
            self.setpointC_var = tk.DoubleVar(value=0.0)
            self.setpointD_var = tk.DoubleVar(value=0.0)

            # Sequence step entry vars
            self.seq_setpointA_var = tk.StringVar(value="")
            self.seq_setpointB_var = tk.StringVar(value="")
            self.seq_setpointC_var = tk.StringVar(value="")
            self.seq_setpointD_var = tk.StringVar(value="")
            self.seq_duration_var = tk.DoubleVar(value=30.0)

            # Sequence control
            self.sequence_running = False
            self.sequence_thread = None
            self.stop_sequence_flag = threading.Event()

            # Data for plotting (time vs flow for A/B/C/D)
            self.start_time = time.time()
            self.time_data = []
            self.A_flow_data = []
            self.B_flow_data = []
            self.C_flow_data = []
            self.D_flow_data = []

            self.create_widgets()
            self.after(self.poll_interval_ms, self.poll_flow)

        # -------------------- UI CREATION --------------------

        def create_widgets(self):
            # Connection frame
            conn_frame = ttk.LabelFrame(self, text="Connection")
            conn_frame.pack(fill="x", padx=10, pady=5)

            ttk.Label(conn_frame, text="Port:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
            self.port_var = tk.StringVar(value="COM10")
            self.port_combo = ttk.Combobox(conn_frame, textvariable=self.port_var, width=10)
            self.port_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")

            if list_ports:
                ports = [p.device for p in list_ports.comports()]
                self.port_combo["values"] = ports

            ttk.Label(conn_frame, text="Baud:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
            self.baud_var = tk.StringVar(value="9600")
            self.baud_entry = ttk.Entry(conn_frame, textvariable=self.baud_var, width=8)
            self.baud_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")

            self.connect_btn = ttk.Button(conn_frame, text="Connect", command=self.on_connect)
            self.connect_btn.grid(row=0, column=4, padx=5, pady=5)

            self.status_label = ttk.Label(conn_frame, text="Status: Disconnected", foreground="red")
            self.status_label.grid(row=0, column=5, padx=5, pady=5, sticky="w")

            # Manual control frame
            manual_frame = ttk.LabelFrame(self, text="Manual Control (4 logical channels)")
            manual_frame.pack(fill="x", padx=10, pady=5)

            # Header row
            headers = ["Logical", "Physical CH", "Setpoint", "Apply", "Flow ON", "Flow OFF"]
            for col, txt in enumerate(headers):
                ttk.Label(manual_frame, text=txt).grid(row=0, column=col, padx=5, pady=3, sticky="w")

            ch_choices = [str(i) for i in range(1, 9)]
            channel_rows = [
                ("A", "Channel A", self.chA_var, self.setpointA_var),
                ("B", "Channel B", self.chB_var, self.setpointB_var),
                ("C", "Channel C", self.chC_var, self.setpointC_var),
                ("D", "Channel D", self.chD_var, self.setpointD_var),
            ]
            for idx, (slot, label, ch_var, sp_var) in enumerate(channel_rows, start=1):
                ttk.Label(manual_frame, text=label).grid(row=idx, column=0, padx=5, pady=3, sticky="w")
                ttk.Combobox(manual_frame, textvariable=ch_var, values=ch_choices, width=5)\
                    .grid(row=idx, column=1, padx=5, pady=3, sticky="w")
                ttk.Entry(manual_frame, textvariable=sp_var, width=10)\
                    .grid(row=idx, column=2, padx=5, pady=3, sticky="w")
                ttk.Button(manual_frame, text="Apply", command=lambda s=slot: self.apply_setpoint_slot(s))\
                    .grid(row=idx, column=3, padx=5, pady=3)
                ttk.Button(manual_frame, text="ON", command=lambda s=slot: self.set_flow_state_slot(s, True))\
                    .grid(row=idx, column=4, padx=5, pady=3)
                ttk.Button(manual_frame, text="OFF", command=lambda s=slot: self.set_flow_state_slot(s, False))\
                    .grid(row=idx, column=5, padx=5, pady=3)

            # Sequence frame

            seq_frame = ttk.LabelFrame(self, text="Sequence (parallel control for Channels A/B/C/D)")
            seq_frame.pack(fill="both", expand=True, padx=10, pady=5)

            columns = ("step", "A", "B", "C", "D", "duration")
            self.seq_tree = ttk.Treeview(seq_frame, columns=columns, show="headings", height=8)
            self.seq_tree.heading("step", text="Step")
            self.seq_tree.heading("A", text="A Setpoint/State")
            self.seq_tree.heading("B", text="B Setpoint/State")
            self.seq_tree.heading("C", text="C Setpoint/State")
            self.seq_tree.heading("D", text="D Setpoint/State")
            self.seq_tree.heading("duration", text="Duration (s)")

            self.seq_tree.column("step", width=50, anchor="center")
            self.seq_tree.column("A", width=120, anchor="center")
            self.seq_tree.column("B", width=120, anchor="center")
            self.seq_tree.column("C", width=120, anchor="center")
            self.seq_tree.column("D", width=120, anchor="center")
            self.seq_tree.column("duration", width=100, anchor="center")

            self.seq_tree.grid(row=0, column=0, columnspan=10, padx=5, pady=5, sticky="nsew")
            seq_frame.rowconfigure(0, weight=1)
            # Configure enough columns so input row aligns fully
            for c in range(10):
                seq_frame.columnconfigure(c, weight=1)

            # Step creation controls - single horizontal row: A, B, C, D, Duration
            # Lay out from the very left so there is no empty gap under the table.
            ttk.Label(seq_frame, text="A SP:").grid(row=1, column=0, padx=5, pady=3, sticky="e")
            ttk.Entry(seq_frame, textvariable=self.seq_setpointA_var, width=8)\
                .grid(row=1, column=1, padx=2, pady=3, sticky="w")

            ttk.Label(seq_frame, text="B SP:").grid(row=1, column=2, padx=5, pady=3, sticky="e")
            ttk.Entry(seq_frame, textvariable=self.seq_setpointB_var, width=8)\
                .grid(row=1, column=3, padx=2, pady=3, sticky="w")

            ttk.Label(seq_frame, text="C SP:").grid(row=1, column=4, padx=5, pady=3, sticky="e")
            ttk.Entry(seq_frame, textvariable=self.seq_setpointC_var, width=8)\
                .grid(row=1, column=5, padx=2, pady=3, sticky="w")

            ttk.Label(seq_frame, text="D SP:").grid(row=1, column=6, padx=5, pady=3, sticky="e")
            ttk.Entry(seq_frame, textvariable=self.seq_setpointD_var, width=8)\
                .grid(row=1, column=7, padx=2, pady=3, sticky="w")

            ttk.Label(seq_frame, text="Duration (s):").grid(row=1, column=8, padx=5, pady=3, sticky="e")
            ttk.Entry(seq_frame, textvariable=self.seq_duration_var, width=8)\
                .grid(row=1, column=9, padx=2, pady=3, sticky="w")

            # Row 2: Control buttons
            self.add_step_btn = ttk.Button(seq_frame, text="Add Step", command=self.add_step)
            self.add_step_btn.grid(row=2, column=0, padx=5, pady=3, sticky="w")

            self.remove_step_btn = ttk.Button(seq_frame, text="Remove Selected", command=self.remove_selected_step)
            self.remove_step_btn.grid(row=2, column=1, padx=5, pady=3, sticky="w")

            self.start_seq_btn = ttk.Button(seq_frame, text="Start Sequence", command=self.start_sequence)
            self.start_seq_btn.grid(row=2, column=3, padx=5, pady=3, sticky="w")

            self.stop_seq_btn = ttk.Button(seq_frame, text="Stop Sequence", command=self.stop_sequence, state="disabled")
            self.stop_seq_btn.grid(row=2, column=4, padx=5, pady=3, sticky="w")

            self.save_seq_btn = ttk.Button(seq_frame, text="Save Sequence", command=self.save_sequence)
            self.save_seq_btn.grid(row=2, column=6, padx=5, pady=3, sticky="w")

            self.load_seq_btn = ttk.Button(seq_frame, text="Load Sequence", command=self.load_sequence)
            self.load_seq_btn.grid(row=2, column=7, padx=5, pady=3, sticky="w")

            # Live plot frame
            plot_frame = ttk.LabelFrame(self, text="Live Plot (Flow vs Time)")
            plot_frame.pack(fill="both", expand=True, padx=10, pady=5)

            # Numeric flow readouts and export buttons above the graph
            info_frame = ttk.Frame(plot_frame)
            info_frame.pack(side="top", fill="x", padx=5, pady=2)

            self.flowA_label = ttk.Label(info_frame, text="A: ---")
            self.flowA_label.pack(side="left", padx=5)
            self.flowB_label = ttk.Label(info_frame, text="B: ---")
            self.flowB_label.pack(side="left", padx=5)
            self.flowC_label = ttk.Label(info_frame, text="C: ---")
            self.flowC_label.pack(side="left", padx=5)
            self.flowD_label = ttk.Label(info_frame, text="D: ---")
            self.flowD_label.pack(side="left", padx=5)

            # Buttons for exporting flow data and saving graph image
            self.export_data_btn = ttk.Button(
                info_frame,
                text="Export Flow Data (Excel)",
                command=self.export_flow_data,
            )
            self.export_data_btn.pack(side="right", padx=5)

            self.save_img_btn = ttk.Button(
                info_frame,
                text="Save Graph Image",
                command=self.save_graph_image,
            )
            self.save_img_btn.pack(side="right", padx=5)

            # Matplotlib figure for live flow plotting
            self.fig = Figure(figsize=(6, 3), dpi=100)
            self.ax = self.fig.add_subplot(111)
            self.ax.set_xlabel("Time (s)")
            self.ax.set_ylabel("Flow (sccm)")
            self.ax.set_ylim(0, 1000)  # Fix Y-axis range from 0 to 1000 sccm
            self.ax.grid(True)

            # Only current flow lines for each logical channel
            (self.line_A_flow,) = self.ax.plot([], [], marker="o", linestyle="-",
                                               color="tab:blue", label="A Flow")
            (self.line_B_flow,) = self.ax.plot([], [], marker="s", linestyle="-",
                                               color="tab:orange", label="B Flow")
            (self.line_C_flow,) = self.ax.plot([], [], marker="^", linestyle="-",
                                               color="tab:green", label="C Flow")
            (self.line_D_flow,) = self.ax.plot([], [], marker="D", linestyle="-",
                                               color="tab:red", label="D Flow")

            self.ax.legend(loc="upper right")

            self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
            self.canvas.get_tk_widget().pack(side="top", fill="both", expand=True)

            # Buttons for exporting flow data and saving graph image
            

        # -------------------- CONNECTION HANDLERS --------------------

        def on_connect(self):
            if not self.controller.connected:
                port = self.port_var.get()
                try:
                    baud = int(self.baud_var.get())
                except ValueError:
                    messagebox.showerror("Error", "Invalid baud rate.")
                    return
                try:
                    self.controller.connect(port, baudrate=baud)
                    self.status_label.config(text=f"Status: Connected ({port})", foreground="green")
                    self.connect_btn.config(text="Disconnect", command=self.on_disconnect)
                except Exception as e:
                    messagebox.showerror("Connection error", str(e))
            else:
                self.on_disconnect()

        def on_disconnect(self):
            try:
                self.controller.disconnect()
            except Exception:
                pass
            self.status_label.config(text="Status: Disconnected", foreground="red")
            self.connect_btn.config(text="Connect", command=self.on_connect)

        # -------------------- MANUAL CONTROL --------------------

        def _get_physical_channel(self, slot):
            if slot == "A":
                return int(self.chA_var.get())
            if slot == "B":
                return int(self.chB_var.get())
            if slot == "C":
                return int(self.chC_var.get())
            if slot == "D":
                return int(self.chD_var.get())
            raise ValueError("Invalid slot")

        def apply_setpoint_slot(self, slot):
            if not self.controller.connected:
                messagebox.showwarning("Not connected", "Connect to the controller first.")
                return
            try:
                if slot == "A":
                    value = float(self.setpointA_var.get())
                elif slot == "B":
                    value = float(self.setpointB_var.get())
                elif slot == "C":
                    value = float(self.setpointC_var.get())
                else:
                    value = float(self.setpointD_var.get())
            except ValueError:
                messagebox.showerror("Error", f"Invalid setpoint for Channel {slot}.")
                return

            ch = self._get_physical_channel(slot)
            try:
                resp = self.controller.set_setpoint(ch, value)
                print(f"[MANUAL] SS{ch} {value} -> {resp}")
                if slot == "A":
                    self.current_setpoint_A = value
                elif slot == "B":
                    self.current_setpoint_B = value
                elif slot == "C":
                    self.current_setpoint_C = value
                else:
                    self.current_setpoint_D = value
            except Exception as e:
                messagebox.showerror("Error", f"Failed to set setpoint for Channel {slot}: {e}")

        def set_flow_state_slot(self, slot, on):
            if not self.controller.connected:
                messagebox.showwarning("Not connected", "Connect to the controller first.")
                return
            ch = self._get_physical_channel(slot)
            try:
                resp = self.controller.set_flow_on_off(ch, on)
                print(f"[MANUAL] SF{ch} {1 if on else 0} -> {resp}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to set flow state for Channel {slot}: {e}")

        # -------------------- PLOTTING HELPERS --------------------

        def _channel_has_activity(self, flow_list):
            """Return True if this channel ever had non-zero flow."""
            for f in flow_list:
                if (not math.isnan(f)) and abs(f) > 1e-6:
                    return True
            return False

        # -------------------- POLLING & PLOTTING --------------------

        def poll_flow(self):
            if self.controller.connected:
                t = time.time() - self.start_time
                self.time_data.append(t)

                # Initialize as NaN
                A_flow_val = math.nan
                B_flow_val = math.nan
                C_flow_val = math.nan
                D_flow_val = math.nan

                # Read Channel A
                try:
                    chA = self._get_physical_channel("A")
                    respA = self.controller.read_flow(chA)
                    if respA:
                        try:
                            A_flow_val = float(respA)
                        except ValueError:
                            A_flow_val = math.nan
                except Exception as e:
                    print(f"[POLL] error reading A: {e}")

                # Read Channel B
                try:
                    chB = self._get_physical_channel("B")
                    respB = self.controller.read_flow(chB)
                    if respB:
                        try:
                            B_flow_val = float(respB)
                        except ValueError:
                            B_flow_val = math.nan
                except Exception as e:
                    print(f"[POLL] error reading B: {e}")

                # Read Channel C
                try:
                    chC = self._get_physical_channel("C")
                    respC = self.controller.read_flow(chC)
                    if respC:
                        try:
                            C_flow_val = float(respC)
                        except ValueError:
                            C_flow_val = math.nan
                except Exception as e:
                    print(f"[POLL] error reading C: {e}")

                # Read Channel D
                try:
                    chD = self._get_physical_channel("D")
                    respD = self.controller.read_flow(chD)
                    if respD:
                        try:
                            D_flow_val = float(respD)
                        except ValueError:
                            D_flow_val = math.nan
                except Exception as e:
                    print(f"[POLL] error reading D: {e}")

                # Append flow data for plot
                self.A_flow_data.append(A_flow_val)
                self.B_flow_data.append(B_flow_val)
                self.C_flow_data.append(C_flow_val)
                self.D_flow_data.append(D_flow_val)

                # Update numeric labels
                if not math.isnan(A_flow_val):
                    self.flowA_label.config(text=f"A: {A_flow_val:.3f} sccm")
                else:
                    self.flowA_label.config(text="A: ---")

                if not math.isnan(B_flow_val):
                    self.flowB_label.config(text=f"B: {B_flow_val:.3f} sccm")
                else:
                    self.flowB_label.config(text="B: ---")

                if not math.isnan(C_flow_val):
                    self.flowC_label.config(text=f"C: {C_flow_val:.3f} sccm")
                else:
                    self.flowC_label.config(text="C: ---")

                if not math.isnan(D_flow_val):
                    self.flowD_label.config(text=f"D: {D_flow_val:.3f} sccm")
                else:
                    self.flowD_label.config(text="D: ---")

                # Update only active channels (show lines only if there's activity)
                if self._channel_has_activity(self.A_flow_data):
                    self.line_A_flow.set_data(self.time_data, self.A_flow_data)
                else:
                    self.line_A_flow.set_data([], [])

                if self._channel_has_activity(self.B_flow_data):
                    self.line_B_flow.set_data(self.time_data, self.B_flow_data)
                else:
                    self.line_B_flow.set_data([], [])

                if self._channel_has_activity(self.C_flow_data):
                    self.line_C_flow.set_data(self.time_data, self.C_flow_data)
                else:
                    self.line_C_flow.set_data([], [])

                if self._channel_has_activity(self.D_flow_data):
                    self.line_D_flow.set_data(self.time_data, self.D_flow_data)
                else:
                    self.line_D_flow.set_data([], [])

                # Rescale and redraw
                self.ax.relim()
                # Autoscale only X; keep Y fixed at 0–1000 sccm
                self.ax.autoscale_view(scalex=True, scaley=False)
                self.canvas.draw_idle()

            self.after(self.poll_interval_ms, self.poll_flow)

        # -------------------- EXPORT / SAVE HELPERS --------------------

        def export_flow_data(self):
            """Export time and flow readings to a CSV file (Excel-compatible)."""
            if not self.time_data:
                messagebox.showwarning("No data", "No flow data to export yet.")
                return

            file_path = filedialog.asksaveasfilename(
                title="Export Flow Data",
                defaultextension=".csv",
                filetypes=[("Excel/CSV Files", "*.csv"), ("All Files", "*.*")],
            )
            if not file_path:
                return

            try:
                with open(file_path, "w", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow([
                        "Time (s)",
                        "A Flow (sccm)",
                        "B Flow (sccm)",
                        "C Flow (sccm)",
                        "D Flow (sccm)",
                    ])
                    # Use len(time_data) as reference
                    n = len(self.time_data)
                    for i in range(n):
                        t = self.time_data[i]
                        a = self.A_flow_data[i] if i < len(self.A_flow_data) else math.nan
                        b = self.B_flow_data[i] if i < len(self.B_flow_data) else math.nan
                        c = self.C_flow_data[i] if i < len(self.C_flow_data) else math.nan
                        d = self.D_flow_data[i] if i < len(self.D_flow_data) else math.nan
                        writer.writerow([t, a, b, c, d])
                messagebox.showinfo("Exported", f"Flow data exported to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export flow data:\n{e}")

        def save_graph_image(self):
            """Save the current matplotlib graph as an image file."""
            file_path = filedialog.asksaveasfilename(
                title="Save Graph Image",
                defaultextension=".png",
                filetypes=[
                    ("PNG Image", "*.png"),
                    ("JPEG Image", "*.jpg;*.jpeg"),
                    ("All Files", "*.*"),
                ],
            )
            if not file_path:
                return

            try:
                self.fig.savefig(file_path, dpi=300, bbox_inches="tight")
                messagebox.showinfo("Saved", f"Graph image saved to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save graph image:\n{e}")

        # -------------------- SEQUENCE MANAGEMENT --------------------

        def add_step(self):
            """Add a sequence step for channels A/B/C/D.

            If a channel setpoint is left blank, it is treated as OFF in that step.
            """
            try:
                duration = float(self.seq_duration_var.get())
            except ValueError:
                messagebox.showerror("Error", "Invalid duration.")
                return

            def build_display(sp_str):
                sp_str = sp_str.strip()
                if sp_str == "":
                    return "-"  # visually show as blank/off
                return sp_str

            A_disp = build_display(self.seq_setpointA_var.get())
            B_disp = build_display(self.seq_setpointB_var.get())
            C_disp = build_display(self.seq_setpointC_var.get())
            D_disp = build_display(self.seq_setpointD_var.get())

            step_index = len(self.seq_tree.get_children()) + 1
            self.seq_tree.insert("", "end", values=(step_index, A_disp, B_disp, C_disp, D_disp, duration))

            # Clear for next
            self.seq_setpointA_var.set("")
            self.seq_setpointB_var.set("")
            self.seq_setpointC_var.set("")
            self.seq_setpointD_var.set("")
            self.seq_duration_var.set(30.0)

        def remove_selected_step(self):
            selected = self.seq_tree.selection()
            for item in selected:
                self.seq_tree.delete(item)
            # Re-number
            for idx, item in enumerate(self.seq_tree.get_children(), start=1):
                vals = list(self.seq_tree.item(item, "values"))
                vals[0] = idx
                self.seq_tree.item(item, values=vals)

        def save_sequence(self):
            if not self.seq_tree.get_children():
                messagebox.showwarning("No steps", "There is no sequence to save.")
                return
            file_path = filedialog.asksaveasfilename(
                title="Save Sequence",
                defaultextension=".csv",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
            )
            if not file_path:
                return
            try:
                with open(file_path, "w", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["step", "A", "B", "C", "D", "duration"])
                    for item in self.seq_tree.get_children():
                        step_no, A_val, B_val, C_val, D_val, dur = self.seq_tree.item(item, "values")
                        writer.writerow([step_no, A_val, B_val, C_val, D_val, dur])
                messagebox.showinfo("Saved", f"Sequence saved to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save sequence:\n{e}")

        def load_sequence(self):
            file_path = filedialog.askopenfilename(
                title="Load Sequence",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
            )
            if not file_path:
                return
            try:
                with open(file_path, "r", newline="") as f:
                    reader = csv.DictReader(f)
                    # Clear existing
                    for item in self.seq_tree.get_children():
                        self.seq_tree.delete(item)
                    step_index = 1
                    for row in reader:
                        A_val = row.get("A", "-")
                        B_val = row.get("B", "-")
                        C_val = row.get("C", "-")
                        D_val = row.get("D", "-")
                        dur = row.get("duration", "0")
                        self.seq_tree.insert("", "end", values=(step_index, A_val, B_val, C_val, D_val, dur))
                        step_index += 1
                messagebox.showinfo("Loaded", f"Sequence loaded from:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load sequence:\n{e}")

        def start_sequence(self):
            if not self.controller.connected:
                messagebox.showwarning("Not connected", "Connect to the controller first.")
                return
            if self.sequence_running:
                messagebox.showinfo("Sequence", "Sequence already running.")
                return

            steps = []
            for item in self.seq_tree.get_children():
                step_no, A_val, B_val, C_val, D_val, dur = self.seq_tree.item(item, "values")
                try:
                    dur = float(dur)
                except ValueError:
                    continue

                setpoints = {}
                on_channels = []
                off_channels = []

                def handle_slot(slot_name, val_str):
                    val_str = str(val_str).strip()
                    if slot_name == "A":
                        ch_str = self.chA_var.get().strip()
                    elif slot_name == "B":
                        ch_str = self.chB_var.get().strip()
                    elif slot_name == "C":
                        ch_str = self.chC_var.get().strip()
                    else:
                        ch_str = self.chD_var.get().strip()
                    if not ch_str:
                        return
                    ch = int(ch_str)

                    # Blank or '-' means this channel should be OFF in this step
                    if val_str == "-" or val_str == "":
                        off_channels.append(ch)
                        return

                    # Otherwise, treat as a numeric setpoint and turn channel ON
                    try:
                        sp_val = float(val_str)
                    except ValueError:
                        return
                    setpoints[ch] = sp_val
                    on_channels.append(ch)

                handle_slot("A", A_val)
                handle_slot("B", B_val)
                handle_slot("C", C_val)
                handle_slot("D", D_val)

                if not setpoints and not off_channels:
                    continue

                steps.append({
                    "duration": dur,
                    "setpoints": setpoints,
                    "on_channels": on_channels,
                    "off_channels": off_channels
                })

            if not steps:
                messagebox.showwarning("No steps", "Add or load at least one step to the sequence.")
                return

            self.sequence_running = True
            self.stop_sequence_flag.clear()
            self.start_seq_btn.config(state="disabled")
            self.stop_seq_btn.config(state="normal")

            self.sequence_thread = threading.Thread(target=self.run_sequence, args=(steps,), daemon=True)
            self.sequence_thread.start()

        def stop_sequence(self):
            if self.sequence_running:
                self.stop_sequence_flag.set()

        def run_sequence(self, steps):
            try:
                for idx, step in enumerate(steps, start=1):
                    # Highlight current step in table
                    try:
                        # Clear previous highlight
                        for item in self.seq_tree.get_children():
                            self.seq_tree.item(item, tags="")
                        # Apply highlight tag
                        current_item = self.seq_tree.get_children()[idx-1]
                        self.seq_tree.item(current_item, tags=("current_step",))
                        self.seq_tree.tag_configure("current_step", background="#ffe08a")
                        # Ensure it is visible
                        self.seq_tree.see(current_item)
                    except Exception:
                        pass
                    if self.stop_sequence_flag.is_set():
                        print("[SEQ] Stopped by user.")
                        break

                    duration = step["duration"]
                    setpoints = step["setpoints"]
                    on_channels = step["on_channels"]
                    off_channels = step["off_channels"]

                    print(f"[SEQ] Step {idx}: setpoints={setpoints}, on={on_channels}, off={off_channels}, dur={duration}s")

                    # Turn OFF required channels first
                    for ch in off_channels:
                        try:
                            resp = self.controller.set_flow_on_off(ch, False)
                            print(f"[SEQ] SF{ch} 0 -> {resp}")
                        except Exception as e:
                            print(f"[SEQ] error turning OFF CH{ch}: {e}")

                    # Apply setpoints
                    for ch, sp_val in setpoints.items():
                        try:
                            resp = self.controller.set_setpoint(ch, sp_val)
                            print(f"[SEQ] SS{ch} {sp_val} -> {resp}")
                            # Update current SP for plotting if matches logical A/B/C/D
                            if ch == int(self.chA_var.get()):
                                self.current_setpoint_A = sp_val
                            if ch == int(self.chB_var.get()):
                                self.current_setpoint_B = sp_val
                            if ch == int(self.chC_var.get()):
                                self.current_setpoint_C = sp_val
                            if ch == int(self.chD_var.get()):
                                self.current_setpoint_D = sp_val
                        except Exception as e:
                            print(f"[SEQ] error setting SP CH{ch}: {e}")

                    # Turn ON channels that should be ON
                    for ch in on_channels:
                        try:
                            resp = self.controller.set_flow_on_off(ch, True)
                            print(f"[SEQ] SF{ch} 1 -> {resp}")
                        except Exception as e:
                            print(f"[SEQ] error turning ON CH{ch}: {e}")

                    # Hold duration
                    start_t = time.time()
                    while time.time() - start_t < duration:
                        if self.stop_sequence_flag.is_set():
                            break
                        time.sleep(0.2)

            except Exception as e:
                print(f"[SEQ] error: {e}")
            finally:
                # SAFETY: always turn all channels OFF when sequence ends or is stopped
                try:
                    print("[SEQ] Turning ALL channels OFF for safety.")
                    self.controller.all_off()
                except Exception as e:
                    print(f"[SEQ] error in all_off: {e}")

                self.sequence_running = False
                self.stop_sequence_flag.clear()
                self.after(0, self._sequence_stopped_ui_update)

        def _sequence_stopped_ui_update(self):
            self.start_seq_btn.config(state="normal")
        
else:

    class App:
        """Fallback App class when tkinter is not available.

        This allows the module to be imported in non-GUI environments.
        Instantiating this class will give a clear error explaining that
        the GUI cannot run without tkinter.
        """

        def __init__(self, *args, **kwargs):
            raise RuntimeError(
                "tkinter is not available in this Python environment; "
                "the GMC1200 GUI cannot be started."
            )


# -------------------- MAIN --------------------

if __name__ == "__main__":
    if tk is None:
        # In a real system you would install/enable tkinter.
        # Here we raise a clear error instead of crashing at import time.
        raise RuntimeError(
            "tkinter is not installed in this Python environment. "
            "Please install tkinter to run the GMC1200 GUI."
        )

    app = App()
    app.mainloop()
