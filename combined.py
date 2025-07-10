import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import time
from datetime import datetime, timedelta
import threading
import queue
import os

# Combined
class CSVProcessorApp:
    def __init__(self, root):
        self.root = root
        self.setup_ui()
        self.setup_variables()
        self.setup_queue()
        
    def setup_ui(self):
        # Window configuration
        self.root.iconphoto(False, tk.PhotoImage(file="icon.png"))
        bg_color = "#1b2b36"
        self.root.configure(background=bg_color)
        
        # Center window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        app_width, app_height = 600, 500
        x = int((screen_width/2) - (app_width/2))
        y = int((screen_height/2) - (app_height/2))
        self.root.geometry(f'{app_width}x{app_height}+{x}+{y-60}')
        self.root.title("Tanium Daily Subscription")
        
        # Header
        header = tk.Label(self.root, text="Tanium Daily Subscription", 
                         font=('Arial', 18, 'bold'), bg='#e1edf5', 
                         fg='#ff003c', width=app_width, height=2)
        header.pack()
        
        # File selection section
        files_label = tk.Label(self.root, text="Select Reports", 
                             font=('Arial', 14, 'bold'), bg=bg_color, 
                             fg='#e1edf5')
        files_label.pack(padx=20, pady=15)
        
        # File buttons
        self.client_status_btn = tk.Button(self.root, text="Last Reg File", 
                                          command=lambda: self.select_file("client_status"),
                                          width=30)
        self.client_status_btn.pack(pady=3)
        
        self.raw_report_btn = tk.Button(self.root, text="RAW Tanium Report", 
                                       command=lambda: self.select_file("raw_report"),
                                       width=30)
        self.raw_report_btn.pack()
        
        self.cmdb_status_btn = tk.Button(self.root, text="CMDB Report", 
                                       command=lambda: self.select_file("cmdb_status"),
                                       width=30)
        self.cmdb_status_btn.pack(pady=3)
        
        self.aws_status_btn = tk.Button(self.root, text="AWS Nimbus Report", 
                                      command=lambda: self.select_file("aws_status"),
                                      width=30)
        self.aws_status_btn.pack()
        
        # Date selection
        footer = tk.Frame(self.root, bg=bg_color)
        footer.pack(pady=15)
        
        date_label = tk.Label(footer, text="Select Date", 
                            font=('Arial', 12, 'bold'), bg=bg_color, 
                            fg='#e1edf5')
        date_label.pack(padx=20)
        
        self.date_var = tk.StringVar()
        self.date_picker = DateEntry(footer, date_pattern="m/d/yyyy", 
                                    textvariable=self.date_var)
        self.date_picker.pack(pady=15)
        
        # Progress bar
        self.progress = ttk.Progressbar(self.root, orient="horizontal", 
                                      length=300, mode="determinate")
        self.progress.pack(pady=10)
        
        # Status label
        self.status_label = tk.Label(self.root, text="", bg=bg_color, 
                                   fg='#e1edf5')
        self.status_label.pack()
        
        # Submit button
        self.submit_btn = tk.Button(self.root, text="Submit", bd="0", 
                                  height=2, width=12, bg='#ff003c', 
                                  fg="#fff", activebackground="#ba0630", 
                                  activeforeground="#fff", 
                                  font=('Arial', 10, 'bold'),
                                  command=self.start_processing)
        self.submit_btn.pack(pady=10)
        
        # Disable submit until files are selected
        self.submit_btn.config(state=tk.DISABLED)
        
    def setup_variables(self):
        self.client_status_file = ""
        self.raw_report_file = ""
        self.cmdb_status_file = ""
        self.aws_status_file = ""
        self.selected_date = self.date_var.get()
        self.date_var.trace("w", self.update_date)
        
    def setup_queue(self):
        self.queue = queue.Queue()
        self.root.after(100, self.process_queue)
        
    def update_date(self, *args):
        self.selected_date = self.date_var.get()
        
    def select_file(self, file_type):
        file_path = filedialog.askopenfilename(title=f"Select {file_type.replace('_', ' ').title()} File", 
                                              filetypes=(("CSV files", "*.csv"), ("All files", "*.*")))
        if file_path:
            if file_type == "client_status":
                self.client_status_file = file_path
                self.client_status_btn.config(text=f"✓ {os.path.basename(file_path)}")
            elif file_type == "raw_report":
                self.raw_report_file = file_path
                self.raw_report_btn.config(text=f"✓ {os.path.basename(file_path)}")
            elif file_type == "cmdb_status":
                self.cmdb_status_file = file_path
                self.cmdb_status_btn.config(text=f"✓ {os.path.basename(file_path)}")
            elif file_type == "aws_status":
                self.aws_status_file = file_path
                self.aws_status_btn.config(text=f"✓ {os.path.basename(file_path)}")
            
            # Enable submit if all files are selected
            if all([self.client_status_file, self.raw_report_file, self.cmdb_status_file, self.aws_status_file]):
                self.submit_btn.config(state=tk.NORMAL)
    
    def start_processing(self):
        # Disable UI during processing
        self.client_status_btn.config(state=tk.DISABLED)
        self.raw_report_btn.config(state=tk.DISABLED)
        self.cmdb_status_btn.config(state=tk.DISABLED)
        self.aws_status_btn.config(state=tk.DISABLED)
        self.date_picker.config(state=tk.DISABLED)
        self.submit_btn.config(state=tk.DISABLED)
        
        # Start processing thread
        processing_thread = threading.Thread(
            target=self.process_csv_files,
            args=(self.client_status_file, self.raw_report_file, 
                 self.cmdb_status_file, self.aws_status_file, 
                 self.selected_date),
            daemon=True
        )
        processing_thread.start()
    
    def process_csv_files(self, client_file, raw_file, cmdb_file, aws_file, selected_date):
        try:
            self.queue.put(("progress", "Starting CSV processing...", 0))
            
            # Your CSV processing logic here
            # Convert selected_date to datetime
            date_chosen = datetime.strptime(selected_date, "%m/%d/%Y")
            date_chosen_older = date_chosen.strftime("%m/%d/%Y")
            date_chosen_newer = (date_chosen + timedelta(days=1)).strftime("%m/%d/%Y")
            
            # Read CSV files into Dataframe
            self.queue.put(("progress", "Reading input files...", 10))
            readInput = pd.read_csv(client_file, usecols=["Host Name","Last Registration"])
            readInput2 = pd.read_csv(raw_file, usecols=["Computer Name","Last Seen", "Operating System"])
            readInput3 = pd.read_csv(cmdb_file, encoding='windows-1252', usecols=["name","install_status", "operational_status"])
            readInput4 = pd.read_csv(aws_file, usecols=["Name","State"])
            
            df = pd.DataFrame(readInput)
            df2 = pd.DataFrame(readInput2)
            df3 = pd.DataFrame(readInput3)
            df4 = pd.DataFrame(readInput4)
            
            # Process raw server report
            self.queue.put(("progress", "Processing raw server report...", 20))
            df2["Date2"] = ""
            for index, cell in df2.iterrows():
                df2.at[index, "Computer Name"] = cell["Computer Name"].split(".")[0]
                df2.at[index, "Date2"] = pd.to_datetime(cell["Last Seen"].split("T")[0])
            
            df2.sort_values(by="Date2", inplace=True, ascending=False)
            df2.drop_duplicates(subset="Computer Name", keep="first", inplace=True)
            
            # Process client status report
            self.queue.put(("progress", "Processing client status report...", 40))
            df["Date"] = ""
            for index, cell in df.iterrows():
                df.at[index, "Host Name"] = cell["Host Name"].split(".")[0]
                df.at[index, "Date"] = pd.to_datetime(cell["Last Registration"].split(",")[0])
            
            df.sort_values(by="Date", inplace=True, ascending=False)
            df.drop_duplicates(subset=['Host Name'], keep="first", inplace=True)
            
            # Filter by date range
            self.queue.put(("progress", "Filtering by date range...", 60))
            final_date_older = datetime.strptime(date_chosen_older, "%m/%d/%Y")
            final_date_newer = datetime.strptime(date_chosen_newer, "%m/%d/%Y")
            
            for index, cell in df.iterrows():
                compared_date = cell["Last Registration"].split(",")[0]
                compared_date_converted = datetime.strptime(compared_date, "%m/%d/%Y")
                if ((compared_date_converted < final_date_older) | (compared_date_converted > final_date_newer)):
                    df.drop(index, axis=0, inplace=True)
            
            # Merge dataframes
            self.queue.put(("progress", "Merging data...", 70))
            df_temp = df.merge(df2, left_on=df["Host Name"].str.lower(), 
                             right_on=df2["Computer Name"].str.lower(), 
                             how="inner")[["Computer Name", "Last Registration", 
                                          "Operating System", "Last Seen"]]
            
            df_temp2 = df_temp.merge(df3, how="left",
                                   left_on=df_temp["Computer Name"].str.lower(),
                                   right_on=df3["name"].str.lower())[["Computer Name", 
                                                                     "Last Registration", 
                                                                     "Operating System",
                                                                     "Last Seen", 
                                                                     "install_status", 
                                                                     "operational_status"]]
            
            df_output = df_temp2.merge(df4, how="left",
                                     left_on=df_temp2["Computer Name"].str.lower(),
                                     right_on=df4["Name"].str.lower())[["Computer Name", 
                                                                       "Last Registration", 
                                                                       "Operating System",
                                                                       "Last Seen", 
                                                                       "install_status", 
                                                                       "operational_status",
                                                                       "State"]]
            
            df_output.drop_duplicates(subset="Computer Name", keep="first", inplace=True)
            df_output = df_output.fillna("N/A")
            
            # Filter out N/A rows
            for index, cell in df_output.iterrows():
                if ((cell["operational_status"] == "N/A") and 
                    (cell["install_status"] == "N/A") and 
                    (cell["State"] == "N/A")):
                    df_output.drop(index, axis=0, inplace=True)
            
            # Add OS Type and count
            self.queue.put(("progress", "Adding OS types and counting...", 80))
            linux_count = windows_count = 0
            cmdb_windows_inactive = cmdb_linux_inactive = 0
            aws_windows_inactive = aws_linux_inactive = 0
            windows_ticket_count = linux_ticket_count = 0
            
            df_output["OS Type"] = ""
            for index, cell in df_output.iterrows():
                if ("AIX" in cell["Operating System"]) or ("Linux" in cell["Operating System"]):
                    df_output.at[index, "OS Type"] = "Linux"
                    linux_count += 1
                else:
                    df_output.at[index, "OS Type"] = "Windows"
                    windows_count += 1
                
                # Count inactive systems
                if ((cell["operational_status"] != "Operational") and (cell["OS Type"] == "Windows")):
                    cmdb_windows_inactive += 1
                if ((cell["operational_status"] != "Operational") and (cell["OS Type"] == "Linux")):
                    cmdb_linux_inactive += 1
                if ((cell["State"] == "stopped") and (cell["OS Type"] == "Windows")):
                    aws_windows_inactive += 1
                if ((cell["State"] == "stopped") and (cell["OS Type"] == "Linux")):
                    aws_linux_inactive += 1
                if ((cell["operational_status"] == "Operational") and (cell["State"] != "stopped") and (cell["OS Type"] == "Windows")):
                    windows_ticket_count += 1
                if ((cell["operational_status"] == "Operational") and (cell["State"] != "stopped") and (cell["OS Type"] == "Linux")):
                    linux_ticket_count += 1
            
            # Save output
            self.queue.put(("progress", "Saving output files...", 90))
            export_path = os.path.dirname(client_file)
            ts = int(time.time())
            
            df_output.to_excel(f'{export_path}/Cleaned_combinedRaw_{ts}.xlsx', index=False)
            df_output.to_excel(f'{export_path}/Cleaned_combined_{ts}.xlsx', index=False)
            
            # Prepare summary
            summary = (
                f"Processing complete!\n\n"
                f"Windows Count: {windows_count}\n"
                f"Linux Count: {linux_count}\n"
                f"CMDB Windows Inactive Count: {cmdb_windows_inactive}\n"
                f"CMDB Linux Inactive Count: {cmdb_linux_inactive}\n"
                f"AWS Windows Inactive Count: {aws_windows_inactive}\n"
                f"AWS Linux Inactive Count: {aws_linux_inactive}\n"
                f"Windows Ticket Count: {windows_ticket_count}\n"
                f"Linux Ticket Count: {linux_ticket_count}\n"
                f"\nFiles saved to: {export_path}"
            )
            
            self.queue.put(("complete", summary, 100))
            
        except Exception as e:
            self.queue.put(("error", f"Error during processing: {str(e)}", 0))
    
    def process_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg[0] == "progress":
                    self.status_label.config(text=msg[1])
                    self.progress["value"] = msg[2]
                elif msg[0] == "complete":
                    self.status_label.config(text=msg[1])
                    self.progress["value"] = msg[2]
                    messagebox.showinfo("Processing Complete", msg[1])
                    self.reset_ui()
                elif msg[0] == "error":
                    self.status_label.config(text=msg[1])
                    messagebox.showerror("Processing Error", msg[1])
                    self.reset_ui()
                self.root.update_idletasks()
        except queue.Empty:
            pass
        self.root.after(100, self.process_queue)
    
    def reset_ui(self):
        self.client_status_btn.config(state=tk.NORMAL)
        self.raw_report_btn.config(state=tk.NORMAL)
        self.cmdb_status_btn.config(state=tk.NORMAL)
        self.aws_status_btn.config(state=tk.NORMAL)
        self.date_picker.config(state=tk.NORMAL)
        self.submit_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = CSVProcessorApp(root)
    root.mainloop()