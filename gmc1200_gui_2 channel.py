import threading
import time
import math
import csv

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
        # Use exactly the same format as the known-working code: command + "\r"
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
        return self.send_cmd(f"SS{ch} {value}")

    def set_flow_on_off(self, ch, on):
        return self.send_cmd(f"SF{ch} {1 if on else 0}")

    def read_flow(self, ch):
        return self.send_cmd(f"RF{ch}")

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
            self.poll_interval_ms = 1000

            # Current setpoints (for reference)
            self.current_setpoint_A = 0.0
            self.current_setpoint_B = 0.0

            # Logical → physical mapping
            self.chA_var = tk.StringVar(value="1")
            self.chB_var = tk.StringVar(value="2")

            # Manual setpoints (direct sccm)
            self.setpointA_var = tk.DoubleVar(value=0.0)
            self.setpointB_var = tk.DoubleVar(value=0.0)

            # Mixture parameters
            self.total_flow_var = tk.DoubleVar(value=1000.0)        # sccm
            self.target_conc_ppm_var = tk.DoubleVar(value=20000.0)  # ppm
            self.carrier_slot_var = tk.StringVar(value="A")        # "A" or "B"
            self.manual_flowppm_var = tk.DoubleVar(value=0.0)       # requested target gas flow in ppm

            # Sequence parameters (in ppm)
            self.seq_flowppm_var = tk.StringVar(value="")
            self.seq_duration_var = tk.DoubleVar(value=30.0)

            # Sequence control
            self.sequence_running = False
            self.sequence_thread = None
            self.stop_sequence_flag = threading.Event()

            # Data for plotting
            self.start_time = time.time()
            self.time_data = []
            self.A_flow_data = []
            self.B_flow_data = []

            self.create_widgets()
            self.after(self.poll_interval_ms, self.poll_flow)

        # -------------------- UI --------------------

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

            # Mixture settings
            mix_frame = ttk.LabelFrame(self, text="Mixture Settings (Total flow + Concentration)")
            mix_frame.pack(fill="x", padx=10, pady=5)

            ttk.Label(mix_frame, text="Total flow (sccm):").grid(row=0, column=0, padx=5, pady=3, sticky="e")
            ttk.Entry(mix_frame, textvariable=self.total_flow_var, width=10).grid(row=0, column=1, padx=5, pady=3, sticky="w")

            ttk.Label(mix_frame, text="Target gas conc (ppm):").grid(row=0, column=2, padx=5, pady=3, sticky="e")
            ttk.Entry(mix_frame, textvariable=self.target_conc_ppm_var, width=10).grid(row=0, column=3, padx=5, pady=3, sticky="w")

            ttk.Label(mix_frame, text="Carrier gas channel:").grid(row=1, column=0, padx=5, pady=3, sticky="e")
            carrier_frame = ttk.Frame(mix_frame)
            carrier_frame.grid(row=1, column=1, columnspan=3, padx=5, pady=3, sticky="w")
            ttk.Radiobutton(carrier_frame, text="A (target = B)", value="A", variable=self.carrier_slot_var).pack(side="left", padx=5)
            ttk.Radiobutton(carrier_frame, text="B (target = A)", value="B", variable=self.carrier_slot_var).pack(side="left", padx=5)

            # Manual control
            manual_frame = ttk.LabelFrame(self, text="Manual Control (2 logical channels)")
            manual_frame.pack(fill="x", padx=10, pady=5)

            headers = ["Logical", "Physical CH", "Setpoint (sccm)", "Apply", "Flow ON", "Flow OFF"]
            for col, txt in enumerate(headers):
                ttk.Label(manual_frame, text=txt).grid(row=0, column=col, padx=5, pady=3, sticky="w")


            ch_choices = [str(i) for i in range(1, 9)]
            channel_rows = [
                ("A", "Channel A", self.chA_var, self.setpointA_var),
                ("B", "Channel B", self.chB_var, self.setpointB_var),
            ]
            for idx, (slot, label, ch_var, sp_var) in enumerate(channel_rows, start=1):
                ttk.Label(manual_frame, text=label).grid(row=idx, column=0, padx=5, pady=3, sticky="w")
                ttk.Combobox(manual_frame, textvariable=ch_var, values=ch_choices, width=5).grid(row=idx, column=1, padx=5, pady=3, sticky="w")
                ttk.Entry(manual_frame, textvariable=sp_var, width=10).grid(row=idx, column=2, padx=5, pady=3, sticky="w")
                ttk.Button(manual_frame, text="Apply", command=lambda s=slot: self.apply_setpoint_slot(s)).grid(row=idx, column=3, padx=5, pady=3)
                ttk.Button(manual_frame, text="ON", command=lambda s=slot: self.set_flow_state_slot(s, True)).grid(row=idx, column=4, padx=5, pady=3)
                ttk.Button(manual_frame, text="OFF", command=lambda s=slot: self.set_flow_state_slot(s, False)).grid(row=idx, column=5, padx=5, pady=3)

            row_base = len(channel_rows) + 1
            ttk.Label(manual_frame, text="Target gas flow (ppm):").grid(row=row_base, column=0, padx=5, pady=3, sticky="e")
            ttk.Entry(manual_frame, textvariable=self.manual_flowppm_var, width=10).grid(row=row_base, column=1, padx=5, pady=3, sticky="w")
            ttk.Button(manual_frame, text="Apply mixture (ppm)", command=self.apply_mixture_manual).grid(row=row_base, column=2, padx=5, pady=3, sticky="w")
            self.mixture_result_label = ttk.Label(manual_frame, text="")
            self.mixture_result_label.grid(row=row_base, column=3, columnspan=4, padx=5, pady=3, sticky="w")

            # Sequence frame
            seq_frame = ttk.LabelFrame(self, text="Sequence (mixture steps for Channels A/B)")
            seq_frame.pack(fill="both", expand=True, padx=10, pady=5)

            columns = ("step", "flow_ppm", "A_sccm", "B_sccm", "duration")
            self.seq_tree = ttk.Treeview(seq_frame, columns=columns, show="headings", height=8)
            self.seq_tree.heading("step", text="Step")
            self.seq_tree.heading("flow_ppm", text="Target flow (ppm)")
            self.seq_tree.heading("A_sccm", text="A flow (sccm)")
            self.seq_tree.heading("B_sccm", text="B flow (sccm)")
            self.seq_tree.heading("duration", text="Duration (s)")

            self.seq_tree.column("step", width=60, anchor="center")
            self.seq_tree.column("flow_ppm", width=120, anchor="center")
            self.seq_tree.column("A_sccm", width=120, anchor="center")
            self.seq_tree.column("B_sccm", width=120, anchor="center")
            self.seq_tree.column("duration", width=100, anchor="center")

            self.seq_tree.grid(row=0, column=0, columnspan=10, padx=5, pady=5, sticky="nsew")
            seq_frame.rowconfigure(0, weight=1)
            for c in range(10):
                seq_frame.columnconfigure(c, weight=1)

            ttk.Label(seq_frame, text="Target flow (ppm):").grid(row=1, column=0, padx=5, pady=3, sticky="e")
            ttk.Entry(seq_frame, textvariable=self.seq_flowppm_var, width=10).grid(row=1, column=1, padx=2, pady=3, sticky="w")

            ttk.Label(seq_frame, text="Duration (s):").grid(row=1, column=2, padx=5, pady=3, sticky="e")
            ttk.Entry(seq_frame, textvariable=self.seq_duration_var, width=10).grid(row=1, column=3, padx=2, pady=3, sticky="w")

            self.add_step_btn = ttk.Button(seq_frame, text="Add Step", command=self.add_step)
            self.add_step_btn.grid(row=1, column=4, padx=5, pady=3, sticky="w")

            self.remove_step_btn = ttk.Button(seq_frame, text="Remove Selected", command=self.remove_selected_step)
            self.remove_step_btn.grid(row=1, column=5, padx=5, pady=3, sticky="w")

            self.start_seq_btn = ttk.Button(seq_frame, text="Start Sequence", command=self.start_sequence)
            self.start_seq_btn.grid(row=1, column=6, padx=5, pady=3, sticky="w")

            self.stop_seq_btn = ttk.Button(seq_frame, text="Stop Sequence", command=self.stop_sequence, state="disabled")
            self.stop_seq_btn.grid(row=1, column=7, padx=5, pady=3, sticky="w")

            self.save_seq_btn = ttk.Button(seq_frame, text="Save Sequence", command=self.save_sequence)
            self.save_seq_btn.grid(row=1, column=8, padx=5, pady=3, sticky="w")

            self.load_seq_btn = ttk.Button(seq_frame, text="Load Sequence", command=self.load_sequence)
            self.load_seq_btn.grid(row=1, column=9, padx=5, pady=3, sticky="w")

            # Plot frame
            plot_frame = ttk.LabelFrame(self, text="Live Plot (Flow vs Time)")
            plot_frame.pack(fill="both", expand=True, padx=10, pady=5)

            info_frame = ttk.Frame(plot_frame)
            info_frame.pack(side="top", fill="x", padx=5, pady=2)

            self.flowA_label = ttk.Label(info_frame, text="A: ---")
            self.flowA_label.pack(side="left", padx=5)
            self.flowB_label = ttk.Label(info_frame, text="B: ---")
            self.flowB_label.pack(side="left", padx=5)

            self.export_data_btn = ttk.Button(info_frame, text="Export Flow Data (Excel)", command=self.export_flow_data)
            self.export_data_btn.pack(side="right", padx=5)

            self.fig = Figure(figsize=(6, 3), dpi=100)
            self.ax = self.fig.add_subplot(111)
            self.ax.set_xlabel("Time (s)")
            self.ax.set_ylabel("Flow (sccm)")
            self.ax.set_ylim(0, 1000)
            self.ax.grid(True)

            (self.line_A_flow,) = self.ax.plot([], [], marker="o", linestyle="-", color="tab:blue", label="A Flow")
            (self.line_B_flow,) = self.ax.plot([], [], marker="s", linestyle="-", color="tab:orange", label="B Flow")

            self.ax.legend(loc="upper right")

            self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
            self.canvas.get_tk_widget().pack(side="top", fill="both", expand=True)

        # -------------------- Connection --------------------

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

        # -------------------- Manual control --------------------

        

        def _get_physical_channel(self, slot):
            if slot == "A":
                return int(self.chA_var.get())
            if slot == "B":
                return int(self.chB_var.get())
            raise ValueError("Invalid slot")

        def apply_setpoint_slot(self, slot):
            if not self.controller.connected:
                messagebox.showwarning("Not connected", "Connect to the controller first.")
                return
            try:
                value = float(self.setpointA_var.get() if slot == "A" else self.setpointB_var.get())
            except ValueError:
                messagebox.showerror("Error", f"Invalid setpoint for Channel {slot}.")
                return

            ch = self._get_physical_channel(slot)
            try:
                resp = self.controller.set_setpoint(ch, value)
                print(f"[MANUAL] SS{ch} {value} -> {resp}")
                if slot == "A":
                    self.current_setpoint_A = value
                else:
                    self.current_setpoint_B = value
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

        # -------------------- Mixture logic --------------------

        def compute_mixture_sccm(self, flow_ppm):
            """Given total flow + target conc, compute A/B sccm for a requested target flow (ppm)."""
            try:
                total = float(self.total_flow_var.get())
                target_ppm = float(self.target_conc_ppm_var.get())
            except ValueError:
                raise ValueError("Invalid total flow or target concentration.")

            if total <= 0 or target_ppm <= 0:
                raise ValueError("Total flow and target concentration must be > 0.")

            F_target = total * (flow_ppm / target_ppm)
            if F_target < 0 or F_target > total:
                raise ValueError("Requested flow ppm is out of range for given total flow and target concentration.")

            F_carrier = total - F_target

            carrier_slot = self.carrier_slot_var.get()
            if carrier_slot == "A":
                A_sccm = F_carrier
                B_sccm = F_target
            else:
                A_sccm = F_target
                B_sccm = F_carrier

            return A_sccm, B_sccm

        def apply_mixture_manual(self):
            """Convert target gas flow in ppm to A/B setpoints *only*.

            This function now:
            - Computes A and B flows from the ppm value
            - Updates the A/B setpoint boxes
            - Updates the result label

            It does NOT send any command to the MFC and does NOT
            turn channels ON/OFF. After using this, you can:
              1) Press the per-channel "Apply" buttons to send
                 the setpoints to the MFC.
              2) Use the per-channel ON/OFF buttons to start/stop flow.
            """
            try:
                flow_ppm = float(self.manual_flowppm_var.get())
            except ValueError:
                messagebox.showerror("Error", "Invalid target gas flow (ppm).")
                return

            try:
                A_sccm, B_sccm = self.compute_mixture_sccm(flow_ppm)
            except ValueError as e:
                messagebox.showerror("Error", str(e))
                return

            # Just update the setpoint entries and the info label.
            # No commands are sent to the controller here.
            self.setpointA_var.set(A_sccm)
            self.setpointB_var.set(B_sccm)

            self.current_setpoint_A = A_sccm
            self.current_setpoint_B = B_sccm
            self.mixture_result_label.config(text=f"A: {A_sccm:.2f} sccm, B: {B_sccm:.2f} sccm")

        # -------------------- Polling & plotting --------------------

        def _channel_has_activity(self, flow_list):
            return any((not math.isnan(f)) and abs(f) > 1e-6 for f in flow_list)

        def poll_flow(self):
            if self.controller.connected:
                t = time.time() - self.start_time
                self.time_data.append(t)

                A_flow_val = math.nan
                B_flow_val = math.nan

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

                self.A_flow_data.append(A_flow_val)
                self.B_flow_data.append(B_flow_val)

                self.flowA_label.config(text=(f"A: {A_flow_val:.3f} sccm" if not math.isnan(A_flow_val) else "A: ---"))
                self.flowB_label.config(text=(f"B: {B_flow_val:.3f} sccm" if not math.isnan(B_flow_val) else "B: ---"))

                if self._channel_has_activity(self.A_flow_data):
                    self.line_A_flow.set_data(self.time_data, self.A_flow_data)
                else:
                    self.line_A_flow.set_data([], [])

                if self._channel_has_activity(self.B_flow_data):
                    self.line_B_flow.set_data(self.time_data, self.B_flow_data)
                else:
                    self.line_B_flow.set_data([], [])

                self.ax.relim()
                self.ax.autoscale_view(scalex=True, scaley=False)
                self.canvas.draw_idle()

            self.after(self.poll_interval_ms, self.poll_flow)

        # -------------------- Export flow data --------------------

        def export_flow_data(self):
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
                    writer.writerow(["Time (s)", "A Flow (sccm)", "B Flow (sccm)"])
                    for i, t in enumerate(self.time_data):
                        a = self.A_flow_data[i] if i < len(self.A_flow_data) else math.nan
                        b = self.B_flow_data[i] if i < len(self.B_flow_data) else math.nan
                        writer.writerow([t, a, b])
                messagebox.showinfo("Exported", f"Flow data exported to:{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export flow data:{e}")

        # -------------------- Sequence management --------------------

        def add_step(self):
            try:
                duration = float(self.seq_duration_var.get())
            except ValueError:
                messagebox.showerror("Error", "Invalid duration.")
                return

            flow_ppm_str = self.seq_flowppm_var.get().strip()
            if not flow_ppm_str:
                messagebox.showerror("Error", "Enter target gas flow (ppm) for this step.")
                return
            try:
                flow_ppm = float(flow_ppm_str)
            except ValueError:
                messagebox.showerror("Error", "Invalid target gas flow (ppm).")
                return

            try:
                A_sccm, B_sccm = self.compute_mixture_sccm(flow_ppm)
            except ValueError as e:
                messagebox.showerror("Error", str(e))
                return

            step_index = len(self.seq_tree.get_children()) + 1
            self.seq_tree.insert("", "end", values=(step_index, flow_ppm, f"{A_sccm:.3f}", f"{B_sccm:.3f}", duration))

            self.seq_flowppm_var.set("")
            self.seq_duration_var.set(30.0)

        def remove_selected_step(self):
            selected = self.seq_tree.selection()
            for item in selected:
                self.seq_tree.delete(item)
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
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            )
            if not file_path:
                return
            try:
                with open(file_path, "w", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["step", "flow_ppm", "A_sccm", "B_sccm", "duration"])
                    for item in self.seq_tree.get_children():
                        step_no, flow_ppm, A_sccm, B_sccm, dur = self.seq_tree.item(item, "values")
                        writer.writerow([step_no, flow_ppm, A_sccm, B_sccm, dur])
                messagebox.showinfo("Saved", f"Sequence saved to:{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save sequence:{e}")

        def load_sequence(self):
            file_path = filedialog.askopenfilename(
                title="Load Sequence",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            )
            if not file_path:
                return
            try:
                with open(file_path, "r", newline="") as f:
                    reader = csv.DictReader(f)
                    for item in self.seq_tree.get_children():
                        self.seq_tree.delete(item)
                    step_index = 1
                    for row in reader:
                        flow_ppm = row.get("flow_ppm", "0")
                        A_sccm = row.get("A_sccm", "0")
                        B_sccm = row.get("B_sccm", "0")
                        dur = row.get("duration", "0")
                        self.seq_tree.insert("", "end", values=(step_index, flow_ppm, A_sccm, B_sccm, dur))
                        step_index += 1
                messagebox.showinfo("Loaded", f"Sequence loaded from:{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load sequence:{e}")

        def _highlight_sequence_step(self, idx):
            for item in self.seq_tree.get_children():
                self.seq_tree.item(item, tags=())
            try:
                current_item = self.seq_tree.get_children()[idx]
            except IndexError:
                return
            self.seq_tree.item(current_item, tags=("current_step",))
            self.seq_tree.tag_configure("current_step", background="#ffe08a")
            self.seq_tree.see(current_item)

        def start_sequence(self):
            if not self.controller.connected:
                messagebox.showwarning("Not connected", "Connect to the controller first.")
                return
            if self.sequence_running:
                messagebox.showinfo("Sequence", "Sequence already running.")
                return

            steps = []
            for item in self.seq_tree.get_children():
                step_no, flow_ppm, A_sccm, B_sccm, dur = self.seq_tree.item(item, "values")
                try:
                    dur_f = float(dur)
                    A_f = float(A_sccm)
                    B_f = float(B_sccm)
                except ValueError:
                    continue

                chA = self._get_physical_channel("A")
                chB = self._get_physical_channel("B")

                setpoints = {}
                on_channels = []
                off_channels = []

                if A_f > 0:
                    setpoints[chA] = A_f
                    on_channels.append(chA)
                else:
                    off_channels.append(chA)

                if B_f > 0:
                    setpoints[chB] = B_f
                    on_channels.append(chB)
                else:
                    off_channels.append(chB)

                if not setpoints and not off_channels:
                    continue

                steps.append({
                    "duration": dur_f,
                    "setpoints": setpoints,
                    "on_channels": on_channels,
                    "off_channels": off_channels,
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
                    self.after(0, self._highlight_sequence_step, idx - 1)

                    if self.stop_sequence_flag.is_set():
                        print("[SEQ] Stopped by user.")
                        break

                    duration = step["duration"]
                    setpoints = step["setpoints"]
                    on_channels = step["on_channels"]
                    off_channels = step["off_channels"]

                    print(f"[SEQ] Step {idx}: setpoints={setpoints}, on={on_channels}, off={off_channels}, dur={duration}s")

                    for ch in off_channels:
                        try:
                            resp = self.controller.set_flow_on_off(ch, False)
                            print(f"[SEQ] SF{ch} 0 -> {resp}")
                        except Exception as e:
                            print(f"[SEQ] error turning OFF CH{ch}: {e}")

                    for ch, sp_val in setpoints.items():
                        try:
                            resp = self.controller.set_setpoint(ch, sp_val)
                            print(f"[SEQ] SS{ch} {sp_val} -> {resp}")
                            if ch == int(self.chA_var.get()):
                                self.current_setpoint_A = sp_val
                            if ch == int(self.chB_var.get()):
                                self.current_setpoint_B = sp_val
                        except Exception as e:
                            print(f"[SEQ] error setting SP CH{ch}: {e}")

                    for ch in on_channels:
                        try:
                            resp = self.controller.set_flow_on_off(ch, True)
                            print(f"[SEQ] SF{ch} 1 -> {resp}")
                        except Exception as e:
                            print(f"[SEQ] error turning ON CH{ch}: {e}")

                    start_t = time.time()
                    while time.time() - start_t < duration:
                        if self.stop_sequence_flag.is_set():
                            break
                        time.sleep(0.2)

            except Exception as e:
                print(f"[SEQ] error: {e}")
            finally:
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
            self.stop_seq_btn.config(state="disabled")
            for item in self.seq_tree.get_children():
                self.seq_tree.item(item, tags=())

else:

    class App:
        def __init__(self, *args, **kwargs):
            raise RuntimeError(
                "tkinter is not available in this Python environment; "
                "the GMC1200 GUI cannot be started."
            )


if __name__ == "__main__":
    if tk is None:
        raise RuntimeError(
            "tkinter is not installed in this Python environment. "
            "Please install tkinter to run the GMC1200 GUI."
        )

    app = App()
    app.mainloop()
