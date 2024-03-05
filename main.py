import sys
import os
import socket
import subprocess
import time
from PyQt5 import QtWidgets, QtCore, QtGui
from openpyxl import Workbook
import ssl

class SpeedTester(QtCore.QThread):
    speed_update = QtCore.pyqtSignal(str, float, float, float, float)

    def __init__(self, host, port, parent=None):
        super().__init__(parent)
        self.host = host
        self.port = port
        self.ping_time = 0
        self.running = False

    def run(self):
        self.running = True
        while self.running:
            self.measure_ping()
            download_speed = self.measure_speed()
            upload_speed = self.measure_upload_speed()
            self.speed_update.emit(self.host, download_speed, upload_speed, 0, self.ping_time)
            time.sleep(1)

    def stop(self):
        print("Stopping SpeedTester...")
        self.running = False
        self.wait()
        print("SpeedTester stopped.")

    def measure_speed(self):
        context = ssl.create_default_context()
        context.verify_mode = ssl.CERT_REQUIRED  
        context.check_hostname = True  
        context.load_default_certs()  

        with socket.create_connection((self.host, self.port)) as sock:
            with context.wrap_socket(sock, server_hostname=self.host) as ssock:
                start_time = time.time()
                host_str = self.host.decode() if isinstance(self.host, bytes) else self.host
                request = 'GET / HTTP/1.1\r\nHost: {}\r\n\r\n'.format(host_str)
                ssock.sendall(request.encode())
                ssock.recv(4096)
                end_time = time.time()

                download_speed = 4096 / (end_time - start_time) / 1024
                return download_speed

    def measure_upload_speed(self):
        context = ssl.create_default_context()
        context.verify_mode = ssl.CERT_REQUIRED
        context.check_hostname = True
        context.load_default_certs()

        with socket.create_connection((self.host, self.port)) as sock:
            with context.wrap_socket(sock, server_hostname=self.host) as ssock:
                start_time = time.time()
                data_to_upload = "This is a test upload string."
                ssock.sendall(data_to_upload.encode())
                response = ssock.recv(1024)
                end_time = time.time()

        upload_speed = len(data_to_upload) / (end_time - start_time) / 1024
        return upload_speed

    def measure_ping(self):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.settimeout(5)  
                start_time = time.time()
                sock.connect((self.host, self.port))
                end_time = time.time()
                self.ping_time = (end_time - start_time) * 1000  
        except (socket.timeout, socket.error):
            self.ping_time = 0

    def measure_datagram_speed(self):
        start_time = time.time()
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.sendto(b'', (self.host, self.port))
            sock.recvfrom(4096)
        end_time = time.time()
        return 4096 / (end_time - start_time) / 1024

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Real-Time Internet Speed Monitor")
        self.setGeometry(100, 100, 283, 524)
        self.setStyleSheet("background-color: #fed183;")

        self.central_widget = QtWidgets.QLabel(self)

        self.speed_labels = {}
        self.speed_testers = []

        hosts_ports = [('www.pesuacademy.com', 443), ('www.facebook.com', 443), ('www.github.com', 443)]

        script_dir = os.path.dirname(os.path.abspath(__file__))
        excel_path = os.path.join(script_dir, "speed_data.xlsx")

        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.sheet.append(["Host ", "Download Speed (Mbps)", "Upload Speed (Mbps)", "Ping Time (ms)"])

        for host, port in hosts_ports:
            speed_label = QtWidgets.QLabel(self)
            speed_label.setGeometry(QtCore.QRect(10, len(self.speed_labels) * 150 + 10, 504, 140))
            font = QtGui.QFont()
            font.setFamily("TIMES NEW ROMAN")
            font.setPointSize(12)
            font.setBold(True)
            speed_label.setFont(font)
            self.speed_labels[host] = speed_label

            speed_tester = SpeedTester(host, port)
            speed_tester.speed_update.connect(self.update_speed_label)
            self.speed_testers.append(speed_tester)

            speed_tester.start()

    @QtCore.pyqtSlot(str, float, float, float, float)
    def update_speed_label(self, host, download_speed, upload_speed, _, ping_time):
        speed_label = self.speed_labels[host]
        speed_label.setText(f"Host: {host}\nDownload Speed: {download_speed:.2f} Mbps\nUpload Speed: {upload_speed:.2f} Mbps\nPing Time: {ping_time:.2f} ms")
        self.write_to_excel(host, download_speed, upload_speed, ping_time)

    def write_to_excel(self, host, download_speed, upload_speed, ping_time):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            excel_path = os.path.join(script_dir, "speed_data.xlsx")
            self.sheet.append([host, download_speed, upload_speed, ping_time])
            self.workbook.save(excel_path)
            print("Data written to Excel successfully.")
        except Exception as e:
            print("An error occurred while writing to Excel:", e)

    def closeEvent(self, event):
        print("Closing application...")
        for tester in self.speed_testers:
            tester.stop()
        super().closeEvent(event)
        print("Application closed.")
        QtCore.QCoreApplication.quit()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
