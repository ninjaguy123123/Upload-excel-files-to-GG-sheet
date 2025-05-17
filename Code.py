import pandas as pd
import numpy as np
import json
import sys
import os
import time
import pickle
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QLineEdit,
                             QPushButton, QVBoxLayout, QWidget, QFileDialog,
                             QTextEdit, QSpinBox, QMessageBox)
from PyQt5.QtCore import QTimer, QThread, pyqtSignal
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import win32com.client as win32
import socket

class WorkerThread(QThread):
    """Thread for handling the automatic upload process."""
    upload_successful = pyqtSignal(str)
    upload_failed = pyqtSignal(str)
    internet_connected = pyqtSignal()
    internet_disconnected = pyqtSignal()

    def __init__(self, folder_path, gsheet_id, service, interval_seconds):
        super().__init__()
        self.folder_path = folder_path
        self.gsheet_id = gsheet_id
        self.service = service
        self.interval_seconds = interval_seconds
        self.running = True
        self.uploaded_files = self.load_uploaded_files()  # Load saved timestamps

    def is_connected(self):
        """Check internet connectivity."""
        try:
            socket.create_connection(("www.google.com", 80), timeout=5)
            return True
        except OSError:
            return False

    def get_existing_google_sheets_data(self):
        """Fetch existing data from Google Sheets."""
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.gsheet_id,
                range='DataGGsheet1!A1'
            ).execute()
            return result.get('values', [])  # Return existing data or empty list
        except Exception as e:
            self.upload_failed.emit(f"Error fetching Google Sheets data: {e}")
            return []

    def get_excel_data(self, file_path):
        """Reads Excel file and extracts data as a list of rows, replacing NaN values."""
        try:
            df = pd.read_excel(file_path, engine="openpyxl", header=None)
            df = df.replace({np.nan: ""})  # Replace NaN with empty string
            return df.values.tolist()
        except Exception as e:
            self.upload_failed.emit(f"Error reading file '{file_path}': {e}")
            return []

    def load_uploaded_files(self):
        """Loads previously uploaded file metadata (filename -> last modified timestamp)."""
        if os.path.exists("uploaded_files.pkl"):
            with open("uploaded_files.pkl", "rb") as f:
                return pickle.load(f)
        return {}

    def save_uploaded_files(self):
        """Saves uploaded file metadata to avoid redundant uploads."""
        with open("uploaded_files.pkl", "wb") as f:
            pickle.dump(self.uploaded_files, f)

    def run(self):
        self.upload_failed.emit("Worker thread started...")
        while self.running:
            if self.is_connected():
                self.internet_connected.emit()
                self.upload_failed.emit("Internet is connected. Scanning folder...")

                if not (self.folder_path and self.gsheet_id and self.service):
                    self.upload_failed.emit("Error: Folder path, Google Sheet ID, or service not initialized.")
                    time.sleep(self.interval_seconds)
                    continue

                try:
                    for filename in os.listdir(self.folder_path):
                        if filename.endswith(('.xlsx', '.xls')):
                            file_path = os.path.join(self.folder_path, filename)
                            file_modified_time = os.path.getmtime(file_path)  # Get last modified timestamp

                            # Check if file has already been uploaded
                            if filename in self.uploaded_files and self.uploaded_files[filename] == file_modified_time:
                                self.upload_failed.emit(f"Bỏ qua '{filename}': Đã tải lên google sheet (No modification).")
                                continue  # Skip to the next file

                            self.upload_failed.emit(f"Kiểm tra file mới: {filename}")

                            # Read Excel data
                            rngData = self.get_excel_data(file_path)
                            if rngData:
                                try:
                                    timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
                                    new_data = [[f"Uploaded on: {timestamp}"]] + rngData

                                    # Read existing Google Sheets data
                                    existing_data = self.get_existing_google_sheets_data()
                                    filtered_data = [row for row in new_data if row not in existing_data]

                                    if filtered_data:
                                        self.service.spreadsheets().values().append(
                                            spreadsheetId=self.gsheet_id,
                                            valueInputOption='USER_ENTERED',
                                            range='DataGGsheet1!A1',
                                            body={'values': filtered_data}
                                        ).execute()
                                        
                                        self.upload_successful.emit(f"File '{filename}' Tải lên vào {timestamp}")

                                        # Save the new modification time
                                        self.uploaded_files[filename] = file_modified_time
                                        self.save_uploaded_files()

                                    else:
                                        self.upload_successful.emit(f"File '{filename}' contains only duplicate data. Skipping upload.")

                                except Exception as e:
                                    self.upload_failed.emit(f"Error sending file '{filename}' to Google Sheets: {e}")

                except FileNotFoundError:
                    self.upload_failed.emit(f"Error: Folder not found: {self.folder_path}")
                except Exception as e:
                    self.upload_failed.emit(f"An unexpected error occurred: {e}")

            else:
                self.internet_disconnected.emit()

            self.upload_failed.emit(f"Đợi {self.interval_seconds} giây để tiếp tục quét...")
            time.sleep(self.interval_seconds)

        self.upload_failed.emit("Worker thread stopped.")




class ExcelToSheetsApp(QMainWindow):
    SETTINGS_FILE = "settings.json"  # File to store saved inputs
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to Google Sheets Uploader")
        self.setGeometry(100, 100, 600, 500)

        self.folder_path = None
        self.gsheet_id = ""
        self.client_secret_file = ""
        self.service = None
        self.upload_interval_minutes = 5  # Default interval

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        # Google Sheet ID Input
        self.gsheet_label = QLabel("Google Sheet ID:")
        self.layout.addWidget(self.gsheet_label)
        self.gsheet_input = QLineEdit()
        self.layout.addWidget(self.gsheet_input)

        # Client Secret File Input
        self.client_secret_label = QLabel("credentials File:")
        self.layout.addWidget(self.client_secret_label)
        self.client_secret_input = QLineEdit()
        self.client_secret_input.setPlaceholderText("Browse or enter path")
        self.layout.addWidget(self.client_secret_input)
        self.browse_client_secret_button = QPushButton("Chọn")
        self.browse_client_secret_button.clicked.connect(self.browse_client_secret)
        self.layout.addWidget(self.browse_client_secret_button)

        # Select Folder
        self.select_folder_label = QLabel("Chọn folder chứa Excel Files:")
        self.layout.addWidget(self.select_folder_label)
        self.folder_path_edit = QLineEdit()
        self.folder_path_edit.setReadOnly(True)
        self.layout.addWidget(self.folder_path_edit)
        self.browse_button = QPushButton("Chọn")
        self.browse_button.clicked.connect(self.browse_folder)
        self.layout.addWidget(self.browse_button)

        # Set Upload Interval
        self.interval_label = QLabel("Thời gian giữa các lần tải tệp (minutes):")
        self.layout.addWidget(self.interval_label)
        self.interval_spinbox = QSpinBox()
        self.interval_spinbox.setMinimum(1)
        self.interval_spinbox.setMaximum(60)
        self.interval_spinbox.setValue(self.upload_interval_minutes)
        self.layout.addWidget(self.interval_spinbox)

        # Notification Board
        self.notification_label = QLabel("Bảng thông báo:")
        self.layout.addWidget(self.notification_label)
        self.notification_panel = QTextEdit()
        self.notification_panel.setReadOnly(True)
        self.layout.addWidget(self.notification_panel)

        # Start Button
        self.start_button = QPushButton("Bắt đầu tự động tải tệp")
        self.start_button.clicked.connect(self.start_upload)
        self.layout.addWidget(self.start_button)

        self.central_widget.setLayout(self.layout)

        # Save Button
        self.save_button = QPushButton("Lưu thiết lập")
        self.save_button.clicked.connect(self.save_settings)
        self.layout.addWidget(self.save_button)

        self.central_widget.setLayout(self.layout)

        # Load previous settings
        self.load_settings()

    def browse_folder(self):
        folder_dialog = QFileDialog()
        folder_path = folder_dialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.folder_path = folder_path
            self.folder_path_edit.setText(self.folder_path)

    def browse_client_secret(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Select Client Secret File", "", "JSON Files (*.json)")
        if file_path:
            self.client_secret_file = file_path
            self.client_secret_input.setText(self.client_secret_file)
    
    def save_settings(self):
        """Save user input to a JSON file."""
        settings = {
            "gsheet_id": self.gsheet_input.text(),
            "client_secret_file": self.client_secret_input.text(),
            "folder_path": self.folder_path_edit.text(),
        }
        with open(self.SETTINGS_FILE, "w") as file:
            json.dump(settings, file)
        print("Settings saved successfully.")

    def load_settings(self):
        """Load user input from a JSON file (if exists)."""
        if os.path.exists(self.SETTINGS_FILE):
            with open(self.SETTINGS_FILE, "r") as file:
                settings = json.load(file)
                self.gsheet_input.setText(settings.get("gsheet_id", ""))
                self.client_secret_input.setText(settings.get("client_secret_file", ""))
                self.folder_path_edit.setText(settings.get("folder_path", ""))

    def create_google_sheets_service(self):
        if not self.client_secret_file:
            self.update_notification("Client Secret File not selected.")
            return
        
        API_SERVICE_NAME = 'sheets'
        API_VERSION = 'v4'
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self.service = self.create_service(self.client_secret_file, API_SERVICE_NAME, API_VERSION, SCOPES)
        
        if self.service:
            self.update_notification("Google Sheets service tạo thành công.")
        else:
            self.update_notification("Thất bại tạo Google Sheets service. kiểm tra credentials.")

    def create_service(self, client_secret_file, api_name, api_version, scopes):
        cred = None
        pickle_file = f'token_{api_name}_{api_version}.pickle'
        
        if os.path.exists(pickle_file):
            with open(pickle_file, 'rb') as token:
                cred = pickle.load(token)
        
        if not cred or not cred.valid:
            if cred and cred.expired and cred.refresh_token:
                cred.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, scopes)
                cred = flow.run_local_server(port=0)
            
            with open(pickle_file, 'wb') as token:
                pickle.dump(cred, token)
        
        try:
            service = build(api_name, api_version, credentials=cred)
            return service
        except Exception as e:
            self.update_notification(f"Unable to connect: {e}")
            return None

    def start_upload(self):
        if hasattr(self, "worker") and self.worker.isRunning():
            # Stop the worker thread
            self.worker.running = False
            self.worker.quit()  # Gracefully stop the thread
            self.worker.wait()
            del self.worker  # Remove reference to the worker
            self.start_button.setText("Bắt đầu tự động tải.")  # Change button text back
            self.update_notification("Dừng tự động tải.")
        else:
            # Start a new worker thread
            self.gsheet_id = self.gsheet_input.text()
            self.client_secret_file = self.client_secret_input.text()

            if not self.gsheet_id or not self.client_secret_file:
                QMessageBox.warning(self, "Warning", "Please enter Google Sheet ID and select the Client Secret File.")
                return

            self.create_google_sheets_service()
            if not self.service:
                self.update_notification("Google Sheets API service is not available. Upload aborted.")
                return

            self.folder_path = self.folder_path_edit.text()
            if not self.folder_path:
                QMessageBox.warning(self, "Warning", "Please select a folder containing Excel files.")
                return

            self.upload_interval_seconds = self.interval_spinbox.value() * 60  # Convert minutes to seconds

            self.worker = WorkerThread(self.folder_path, self.gsheet_id, self.service, self.upload_interval_seconds)

            # Connect signals
            self.worker.upload_successful.connect(self.update_notification)
            self.worker.upload_failed.connect(self.update_notification)
            self.worker.internet_connected.connect(lambda: self.update_notification("Internet đã kết nối"))
            self.worker.internet_disconnected.connect(lambda: self.update_notification("Internet mất kết nối"))

            self.worker.start()
            self.start_button.setText("Dừng tự động tải.")  # Change button text
            self.update_notification("Bắt đầu tự động tải.")




    def update_notification(self, message):
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        log_message = f"[{timestamp}] {message}"
        self.notification_panel.append(log_message)
        self.notification_panel.ensureCursorVisible()

        # Save to file
        #with open("notifications.txt", "a", encoding="utf-8") as log_file:
            #log_file.write(log_message + "\n")
            #log_file.flush()  # Ensures data is written immediately


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = ExcelToSheetsApp()
    main_window.show()
    sys.exit(app.exec_())