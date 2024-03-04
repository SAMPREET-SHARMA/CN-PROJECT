README FILE

### Real-Time Internet Speed Monitor

This Python application offers a real-time monitoring solution for internet speed, providing insights into download speed, upload speed, and ping time for multiple hosts. The application features a user-friendly graphical user interface (GUI) and automatically logs the collected data into an Excel file for further analysis.

### How to Run the Code

#### Install Dependencies:

Ensure Python is installed on your system. You can download and install Python from the official website: [Python Downloads](https://www.python.org/downloads/).

Once Python is installed, you can install the required Python packages using pip, a package manager for Python. Open your terminal or command prompt and execute the following command:

(_**bash
pip install PyQt5 openpyxl**_)


#### Clone the Repository:

Clone this GitHub repository to your local machine. You can do this by navigating to the directory where you want to store the repository and executing the following command in your terminal or command prompt:

_**bash
git clone https://github.com/your-username/real-time-speed-monitor.git**_


#### Run the Application:

Navigate to the directory containing the cloned repository. You can do this by using the cd command followed by the path to the directory. Once you're in the correct directory, execute the following command in your terminal or command prompt:

_**bash
python main.py**_


### Overview of the Code

- **main.py:** This is the main Python script responsible for initializing the application and GUI. It creates instances of the SpeedTester class for each host and starts the monitoring threads.

- **SpeedTester Class:**
  - This class extends QtCore.QThread, allowing for concurrent execution of speed tests.
  - It measures download speed, upload speed, and ping time using socket connections.
  - Results are emitted through the speed_update signal.

- **MainWindow Class:**
  - This class represents the main application window.
  - It initializes the GUI, including labels for displaying speed data.
  - Hosts and ports to monitor are defined within this class.
  - Speed testers are created and started, and their signals are connected to update GUI labels.

- *Excel Data Logging:*
  - The application writes speed data to an Excel file named speed_data.xlsx.
  - Each row contains information about a specific host, including download speed, upload speed, and ping time.

### Closing the Application

- The application can be closed by clicking the close button on the window.
- Upon closure, the application stops all speed testers and saves the Excel file.
- The application can also be closed programmatically using the closeEvent method.

### Note

- Ensure you have an active internet connection to monitor speeds effectively.
- You can customize the code to monitor additional hosts by modifying the hosts_ports list in the MainWindow class.
- Contributions to the code, such as adding new features or improving existing ones, are welcome. Feel free to fork the repository and submit pull requests.
