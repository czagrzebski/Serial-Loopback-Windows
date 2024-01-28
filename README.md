<div align="center">
 <h3>Serial Port Loopback for Windows </h3>
 <img src="logo.png" width="20%">
</div>

## Introduction
The Serial Port Loopback application is designed for Windows and provides an automated interface for serial devices. It continuously monitors USB ports for any connection or disconnection events. When a supported serial device is connected, the application automatically establishes a serial connection, creating a loopback interface where any data received on the serial port is sent back to the device. This can be particularly useful for testing serial communication or for debugging purposes.

## Features
- **Auto-detection of Device Connection/Disconnection**: Leverages Windows Management Instrumentation (WMI) to detect when a USB device is plugged in or unplugged.
- **Serial Communication**: Utilizes pySerial to open a serial connection to the device and perform data loopback.
- **Customizable Device Support**: Allows adding and removing of supported devices based on Vendor ID (VID) and Product ID (PID).
- **Flexible Settings**: Adjustable baud rate and support for multiple devices.
- **User-friendly Interface**: Provides a GUI to interact with the application, view connected devices, and adjust settings.

## Setup

### Prerequisites
Ensure you have the following installed:
- Python 3.x
- pySerial: `pip install pyserial`
- pywin32: `pip install pywin32`
- PyQt5: `pip install pyqt5`

### Installation (Method 1)
1. **Clone the repository or download the source code.**
   ```bash
   git clone https://github.com/czagrzebski/Serial-Loopback-Windows.git
   cd Serial-Loopback-Windows
   ```
2. **Install the required packages.**
   ```bash
   pip install -r requirements.txt
   ```

### Installation (Method 2)
1. Download x86 or x64 executable from the [releases](https://github.com/czagrzebski/Serial-Loopback-Windows/releases) page.
2. Run the executable.

## Usage
1. **Run the application using command line or by double-clicking the executable.**
   ```bash
   python main.py
   ```
2. **Adjust the settings as needed**
    - **Baud Rate**: The baud rate to use for the serial connection.
    - **Device List**: The list of supported devices. To add a device, enter the VID and PID in the text boxes and click the "Add" button. To remove a device, select it from the list and click the "Remove" button.
3. **Connect a supported USB serial device.**
    - The application automatically detects connected devices. If the device's VID and PID are in the list of supported devices, it establishes a serial connection.
    - Once connected, the device is added to the list of connected devices. A thread is created to monitor the serial port for incoming data. Any data received on the serial port is sent back to the device.
4. **Disconnect the device.**
    - The application automatically detects when a device is disconnected and closes the serial connection.
    - To manually disconnect a device, select it from the list and click the "Disconnect" button.
    - To manually disconnect all devices, click the "Disconnect All" button.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
