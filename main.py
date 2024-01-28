"""
USB Serial Port Loopback for Windows

This application provides a loopback interface for serial devices on Windows Machines.

Author: Creed Zagrzebski (czagrzebski@gmail.com)

"""

import wmi
import serial
import serial.tools.list_ports
import threading
import time
import sys
import pythoncom
from PyQt5.QtWidgets import QApplication, QMainWindow, QListWidget, QVBoxLayout, QWidget, QLabel, QDialog, QPushButton, QLineEdit, QMenu, QCheckBox
from PyQt5.QtCore import pyqtSignal, pyqtSlot, QThread
from PyQt5.QtCore import Qt
import json

# Global variables for settings
supported_devices = []  # List of supported devices in the format 'VID:PID=xxxx:xxxx'
BAUD_RATE = 115200
AUTO_CONNECT = True
PARITY = serial.PARITY_NONE
SETTINGS_FILE = 'settings.json'

def get_serial_ports():
    # returns a dictionary of serial ports in the format {com_port: hwid} along with the name of the serial port
    return {port.device: [port.hwid, port.manufacturer, port.description] for port in serial.tools.list_ports.comports()}

def is_supported_device(device_hwid):
    for device in supported_devices:
        if device in device_hwid:
            return True
    return False

class SerialDevice:
    def __init__(self, com_port, vid, pid, serial_connection=None):
        self.com_port = com_port
        self.vid = vid
        self.pid = pid
        self.serial_connection = serial_connection

    def __str__(self):
        return f"{self.com_port} (VID: {self.vid}, PID: {self.pid})"

class SerialReadThread(threading.Thread):
    def __init__(self, serial_devices, serial_device, disconnect_callback=None):
        super().__init__(daemon=True)
        self.serial_devices = serial_devices
        self.serial_device = serial_device
        self.disconnect_callback = disconnect_callback

    def run(self):
        try:
            while self.serial_device.serial_connection.is_open:
                # this is more efficient than busy waiting, significantly reduces CPU usage
                data = self.serial_device.serial_connection.read(self.serial_device.serial_connection.in_waiting)
                if data:  # If any data is received
                    self.serial_device.serial_connection.write(data)
        except serial.SerialException as e:
            print(f"Serial exception on {self.serial_device.com_port}: {e}")
        finally:
            # remove device from list of serial devices
            self.serial_devices.remove(self.serial_device)
            # emit signal to remove device from list
            if self.disconnect_callback:
                self.disconnect_callback.emit(self.serial_device.com_port)
            print(f"Closed connection on {self.serial_device.com_port}")

class DeviceMonitorThread(QThread):
    device_connected_signal = pyqtSignal(str, str, str, str, str)
    device_disconnected_signal = pyqtSignal(str)

    def __init__(self, serial_devices):
        super().__init__()
        self.serial_devices = serial_devices
        self.running = True

    def run(self):
        
        # there is some magic going on here with COM threading and WMI, but it works so ¯\_(ツ)_/¯
        pythoncom.CoInitialize()
        c = wmi.WMI()
        
        # YOU SCREAM, I SCREAM, WE ALL SCREAM FOR WQL!!
        # honestly though, this is just cursed SQL that hurts my brain
        
        # WITHIN 2 means that the event will be triggered if the device is connected or disconnected within 2 seconds
        # TargetInstance ISA 'Win32_USBhub' means that the event will be triggered if the device is a USB hub
        # TargetInstance ISA 'Win32_SerialPort' means that the event will be triggered if the device is a serial port
        # The ISA operator is used in the WHERE clause of data and event queries to test embedded objects for a class hierarchy.
        device_connected_wql = "SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBhub' or TargetInstance ISA 'Win32_SerialPort'"
        device_disconnected_wql = "SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBhub' or TargetInstance ISA 'Win32_SerialPort'"

        # watch for changes in USB devices (connected and disconnected)
        connected_watcher = c.watch_for(raw_wql=device_connected_wql)
        disconnected_watcher = c.watch_for(raw_wql=device_disconnected_wql)
        
        # detect devices on startup
        if AUTO_CONNECT:
            self.detect_devices()

        while self.running:
            if AUTO_CONNECT:
                try:
                    connected = connected_watcher(timeout_ms=10)
                    if connected:
                        print("Change in USB devices detected. Checking for new devices...")
                        time.sleep(3) # this is a hacky way to wait for the device to be fully connected but it works ¯\_(ツ)_/¯
                        self.detect_devices()
                except wmi.x_wmi_timed_out:
                    pass
                try:
                    disconnected = disconnected_watcher(timeout_ms=10)
                    if disconnected:
                        print("Change in USB devices detected. Checking for new devices...")
                        self.detect_devices()
                except wmi.x_wmi_timed_out:
                    pass
        pythoncom.CoUninitialize()

    def detect_devices(self):
        print("Detecting devices...")
        new_ports = get_serial_ports()
        existing_com_ports = {device.com_port for device in self.serial_devices}

        # Detect new devices
        for com_port, desc in new_ports.items():
            # if the com port is not already in the list of serial devices and the device is supported (VID:PID is in hwid)
            if com_port not in existing_com_ports and is_supported_device(desc[0]):
                # get vid and pid from hwid
                vid_pid = desc[0].split('VID:PID=')[1].split(' ')[0] if 'VID:PID=' in desc[0] else 'Unknown'
                vid, pid = vid_pid.split(':')
                try:
                    print(f"Opening new serial connection to {desc[2]} ({desc[1]})")
                    ser = serial.Serial(com_port, BAUD_RATE, parity=PARITY, timeout=0.5)
                    new_device = SerialDevice(com_port, vid, pid, ser)
                    self.serial_devices.append(new_device)
                    t = SerialReadThread(self.serial_devices, new_device, self.device_disconnected_signal)
                    t.daemon = True # thread dies when main thread (only non-daemon thread) exits. don't want a bunch of zombie threads
                    t.start() # start the thread
                    # emit signal to add device to list
                    self.device_connected_signal.emit(com_port, vid, pid, desc[1], desc[2])
                except serial.SerialException as e:
                    print(f"Error opening serial connection on {com_port}: {e}")

        # Disconnect removed devices
        for device in self.serial_devices[:]:
            if device.com_port not in new_ports:
                if device.serial_connection and device.serial_connection.is_open:
                    device.serial_connection.close()
                self.serial_devices.remove(device)
                self.device_disconnected_signal.emit(device.com_port)

    def stop(self):
        self.running = False
        
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.serial_devices = []
        self.initUI()
        self.device_monitor_thread = DeviceMonitorThread(self.serial_devices)
        self.device_monitor_thread.device_connected_signal.connect(self.addDevice)
        self.device_monitor_thread.device_disconnected_signal.connect(self.removeDevice)
        self.device_monitor_thread.start()

    def initUI(self):
        self.setWindowTitle("Serial Loopback for Windows")
        self.setGeometry(100, 100, 800, 600)
        # make window non resizable
        self.setFixedSize(self.size())
        layout = QVBoxLayout()
        
        layout.addWidget(QLabel("Active USB Loopback Devices"))

        self.deviceList = QListWidget()
        layout.addWidget(self.deviceList)
                
        self.detectButton = QPushButton("Detect Loopback Devices", self)
        self.detectButton.clicked.connect(lambda: self.device_monitor_thread.detect_devices())
        layout.addWidget(self.detectButton)
        
        # Disconnect all devices button
        self.disconnectButton = QPushButton("Disconnect All Loopback Devices", self)
        self.disconnectButton.clicked.connect(self.disconnectAll)
        layout.addWidget(self.disconnectButton)
   
        
        self.settingsButton = QPushButton("Settings", self)
        self.settingsButton.clicked.connect(self.openSettingsDialog)
        layout.addWidget(self.settingsButton)
        
        self.aboutButton = QPushButton("About", self)
        self.aboutButton.clicked.connect(self.openAboutDialog)
        layout.addWidget(self.aboutButton)
        
        self.deviceList.setContextMenuPolicy(Qt.CustomContextMenu)
        self.deviceList.customContextMenuRequested.connect(self.openDeviceMenu)

        centralWidget = QWidget()
        centralWidget.setLayout(layout)
        self.setCentralWidget(centralWidget)
        
    def disconnectAll(self):
        for device in self.serial_devices[:]:
            if device.serial_connection and device.serial_connection.is_open:
                try:
                    device.serial_connection.close()
                except serial.SerialException as e:
                    print(f"Error closing serial connection on {device.com_port}: {e}")
                    
    def openDeviceMenu(self, position):
        menu = QMenu()
        disconnect_action = menu.addAction("Disconnect")
        action = menu.exec_(self.deviceList.viewport().mapToGlobal(position))
        
        # Disconnect device
        if action == disconnect_action:
            selected_item = self.deviceList.currentItem()
            if selected_item:
                selected_device_str = selected_item.text()
                self.disconnectDevice(selected_device_str.split(' ')[0])  # Extracting the COM port

    def disconnectDevice(self, com_port):
        for device in self.serial_devices:
            if device.com_port == com_port:
                if device.serial_connection and device.serial_connection.is_open:
                    device.serial_connection.close()
                break

    def openSettingsDialog(self):
        dialog = SettingsDialog(self)
        dialog.exec_()

    @pyqtSlot(str, str, str, str, str)
    def addDevice(self, com_port, vid, pid, manufacturer, description):
        device_info = f"{com_port} (Device: {description.split(' (COM')[0]}, Manufacturer: {manufacturer}, VID: {vid}, PID: {pid})"
        self.deviceList.addItem(device_info)

    @pyqtSlot(str)
    def removeDevice(self, com_port):
        for i in range(self.deviceList.count()):
            item = self.deviceList.item(i)
            if item.text().startswith(com_port + ' '):
                self.deviceList.takeItem(i)
                break

    def closeEvent(self, event):
        self.device_monitor_thread.stop()
        event.accept()
        
    def openAboutDialog(self):
        dialog = AboutDialog(self)
        dialog.exec_()
        
class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super(SettingsDialog, self).__init__(parent)
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Settings")
        self.layout = QVBoxLayout(self)

        # VID and PID input
        self.vidLabel = QLabel("VID:")
        self.vidInput = QLineEdit(self)
        self.pidLabel = QLabel("PID:")
        self.pidInput = QLineEdit(self)
        self.addDeviceButton = QPushButton("Add Device", self)
        self.addDeviceButton.clicked.connect(self.addDevice)
        self.layout.addWidget(self.vidLabel)
        self.layout.addWidget(self.vidInput)
        self.layout.addWidget(self.pidLabel)
        self.layout.addWidget(self.pidInput)
        self.layout.addWidget(self.addDeviceButton)

        # Device list
        self.deviceList = QListWidget(self)
        self.layout.addWidget(self.deviceList)
        self.removeButton = QPushButton("Remove Selected", self)
        self.removeButton.clicked.connect(self.removeSelected)
        self.layout.addWidget(self.removeButton)

        # Baud rate input
        self.baudLabel = QLabel("Baud Rate:")
        self.baudInput = QLineEdit(self)
        self.baudInput.setText(str(BAUD_RATE))
        self.layout.addWidget(self.baudLabel)
        self.layout.addWidget(self.baudInput)
        
        self.autoConnect = QLabel("Automatically connect to devices on USB event:")
        self.autoConnectCheckbox = QCheckBox(self)
        self.autoConnectCheckbox.setChecked(AUTO_CONNECT)
        self.layout.addWidget(self.autoConnect)
        self.layout.addWidget(self.autoConnectCheckbox)
        
        # Save button
        self.saveButton = QPushButton("Save and Exit", self)
        self.saveButton.clicked.connect(self.saveSettings)
        self.layout.addWidget(self.saveButton)

        self.loadSettings()

    def addDevice(self):
        vid = self.vidInput.text()
        pid = self.pidInput.text()
        if vid and pid:
            device = f'VID:PID={vid}:{pid}'
            supported_devices.append(device)
            self.deviceList.addItem(device)

    def removeSelected(self):
        listItems = self.deviceList.selectedItems()
        if not listItems: return
        for item in listItems:
            supported_devices.remove(item.text())
            self.deviceList.takeItem(self.deviceList.row(item))

    def saveSettings(self):
        global BAUD_RATE, AUTO_CONNECT
        baud_rate = self.baudInput.text()
        auto_connect = self.autoConnectCheckbox.isChecked()
        if baud_rate.isdigit():
            BAUD_RATE = int(baud_rate)
        AUTO_CONNECT = auto_connect
        saveSettings()
        self.close()

    def loadSettings(self):
        self.deviceList.clear()
        for device in supported_devices:
            self.deviceList.addItem(device)
            
class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super(AboutDialog, self).__init__(parent)
        self.initUI()

    def initUI(self):
        self.setWindowTitle("About")
        self.layout = QVBoxLayout(self)

        # gotta plug my github lol
        github_url = "https://github.com/czagrzebski"  
        self.aboutLabel = QLabel(f"Serial Loopback for Windows<br>Version 1.0.3<br><br>" +
                                 "This application provides a loopback interface for serial devices on Windows Machines.<br><br>" +
                                 "Developed by Creed Zagrzebski.<br><br>" +
                                 "Licensed under the MIT License.<br><br>" +
                                 f"<a href='{github_url}'>{github_url}", self)
        self.aboutLabel.setWordWrap(True)
        self.aboutLabel.setOpenExternalLinks(True)  # Allow label to open links in a web browser
        self.layout.addWidget(self.aboutLabel)

        # Close button
        self.closeButton = QPushButton("Close", self)
        self.closeButton.clicked.connect(self.close)
        self.layout.addWidget(self.closeButton)
            
def saveSettings():
    settings = {
        'supported_devices': supported_devices,
        'baud_rate': BAUD_RATE,
        'auto_connect': AUTO_CONNECT
    }
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings, f)

def loadSettings():
    global BAUD_RATE, AUTO_CONNECT, supported_devices
    try:
        with open(SETTINGS_FILE, 'r') as f:
            settings = json.load(f)
            supported_devices = settings.get('supported_devices', [])
            BAUD_RATE = settings.get('baud_rate', 115200)
            AUTO_CONNECT = settings.get('auto_connect', True)
    except FileNotFoundError:
        pass


if __name__ == '__main__':
    print("========================================")
    print("USB Serial Port Loopback for Windows")
    print("By Creed Zagrzebski")
    print("========================================")
    loadSettings()
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())