import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import win32com.client

class CanoeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CANoe Configuration Tool")
        self.selected_config = tk.StringVar(value="No configuration selected")
        self.is_running = False
        self.canoe_app = None

        self.create_widgets()
        self.resizable(True, True)  # Allowing window resizing

    def create_widgets(self):
        # Configuration Selection Section
        config_frame = tk.Frame(self, padx=20, pady=20, bd=2, relief=tk.RAISED)
        config_frame.pack(fill=tk.X)

        tk.Label(config_frame, text="CANoe Configuration:", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w", padx=10)
        tk.Entry(config_frame, textvariable=self.selected_config, width=60, font=("Arial", 11)).grid(row=1, column=0, padx=10, pady=10)
        tk.Button(config_frame, text="Select File", command=self.select_configuration, font=("Arial", 11)).grid(row=1, column=1, padx=10, pady=10)

        # Measurement Control Section
        control_frame = tk.Frame(self, padx=20, pady=20, bd=2, relief=tk.RAISED)
        control_frame.pack(fill=tk.X)

        self.run_button = tk.Button(control_frame, text="RUN", command=self.run_measurement, state=tk.DISABLED, font=("Arial", 12))
        self.run_button.pack(side=tk.LEFT, padx=10)

        self.stop_measurement_button = tk.Button(control_frame, text="Stop Measurement", command=self.stop_measurement, state=tk.DISABLED, font=("Arial", 12))
        self.stop_measurement_button.pack(side=tk.LEFT, padx=10)

        self.close_canoe_button = tk.Button(control_frame, text="Close CANoe", command=self.close_canoe, state=tk.DISABLED, font=("Arial", 12))
        self.close_canoe_button.pack(side=tk.LEFT, padx=10)

        # Log Section
        log_frame = tk.Frame(self, padx=20, pady=20, bd=2, relief=tk.RAISED)
        log_frame.pack(fill=tk.BOTH, expand=True)

        log_label = tk.Label(log_frame, text="Log", font=("Arial", 12, "bold"))
        log_label.pack(side=tk.TOP, padx=10, pady=10)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, font=("Arial", 11), height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Adding a separator
        separator = tk.Frame(self, height=2, bd=1, relief=tk.SUNKEN)
        separator.pack(fill=tk.X, padx=5, pady=5)

        # Adding status bar
        self.status_bar = tk.Label(self, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 10))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def select_configuration(self):
        filename = filedialog.askopenfilename(filetypes=[("Configuration files", "*.cfg")])
        if filename:
            self.selected_config.set(filename)
            self.log_action(f"Selected configuration: {filename}")
            self.run_button.config(state=tk.NORMAL)

    def run_measurement(self):
        config = self.selected_config.get()
        if config:
            try:
                self.canoe_app = win32com.client.Dispatch("CANoe.Application")
                self.canoe_app.Open(config)
                self.canoe_app.Measurement.Start()

                self.is_running = True
                self.run_button.config(state=tk.DISABLED)
                self.stop_measurement_button.config(state=tk.NORMAL)
                self.close_canoe_button.config(state=tk.NORMAL)
                self.log_action("CANoe opened and measurement started")
                self.status_bar.config(text="Measurement Running")
            except Exception as e:
                self.log_action(f"Error: Failed to open CANoe: {str(e)}")
        else:
            self.log_action("Error: Please select a configuration")

    def stop_measurement(self):
        if self.is_running:
            try:
                measurement = self.canoe_app.Measurement
                configuration = self.canoe_app.Configuration

                measurement.Stop()
                configuration.Save()

                self.is_running = False
                self.run_button.config(state=tk.NORMAL)
                self.stop_measurement_button.config(state=tk.DISABLED)
                self.close_canoe_button.config(state=tk.NORMAL)
                self.log_action("Measurement stopped")
                self.status_bar.config(text="Measurement Stopped")
            except Exception as e:
                self.log_action(f"Error: Failed to stop CANoe measurement: {str(e)}")
        else:
            self.log_action("Info: No measurement is currently running")

    def close_canoe(self):
        if self.canoe_app is not None:
            try:
                self.canoe_app.Quit()
                self.canoe_app.Close()
                self.log_action("CANoe closed")
                self.canoe_app = None
                self.close_canoe_button.config(state=tk.DISABLED)
                self.status_bar.config(text="CANoe Closed")
            except Exception as e:
                self.log_action(f"Error: Failed to close CANoe: {str(e)}")
        else:
            self.log_action("Info: CANoe is not running")

    def log_action(self, action):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {action}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)  # Scroll to the end of the log
        self.status_bar.config(text=action)

if __name__ == "__main__":
    app = CanoeApp()
    app.mainloop()
