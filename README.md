# `touchpoint-meja`

![Static Badge](https://img.shields.io/badge/goal-Transform_Azure_SQL_Data_to_Excel-purple)
<br />
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

<p class="align center">
<h4><code>touchpoint-meja</code> is a Python script for transforming touchpoint data for the <code>NCR-MothersEntJA</code> database to an Excel Workbook.</h4>
</p>

[What and Why](#whatandwhy) •
[Quickstart](#quickstart) •
[Running the Script](#running-the-script)

</div>

## Navigation

- [Introduction Videos](#introduction-videos)
- [What and Why](#what-and-why)
- [Quickstart](#quickstart)
  - [Required Python Version](#required-python-version-)
  - [Using Command Prompt or PowerShell](#using-command-prompt-or-powershell)
  - [If Python is Not Installed](#if-python-is-not-installed)
  - [Troubleshooting](#troubleshooting)
  - [Downloading the script](#downloading-the-script)
  - [Installing the Packages](#installing-the-packages)
- [Running the Script](#running-the-script)

<br />

## What and why

The purpose of this Python scrip is to 

## Quickstart

### Required Python Version 
Ensure you have at least python3.6 installed on you operating system. Otherwise, when you attempt to run the pip install commands, the project will fail to build due to certain dependencies. 
To verify if Python is installed on Windows, you can use the Command Prompt (cmd) or PowerShell. Here's a step-by-step guide:

### Using Command Prompt or PowerShell

1. **Open Command Prompt or PowerShell**:
   - Press `Win + R`, type `cmd`, and press `Enter` to open Command Prompt.
   - Alternatively, press `Win + X` and select `Windows PowerShell` or `Windows Terminal` if you're using a more recent version of Windows.

2. **Check Python Installation**:
   - Type `python --version` and press `Enter`.
   - Alternatively, you can type `python3 --version` and press `Enter`.

   If Python is installed, you will see the version number, such as:

   ```sh
   Python 3.x.x
   ```

3. **Check Python Path**:
   - To ensure that Python is added to your system's PATH, you can type `where python` and press `Enter`.

   This command will show you the path where Python is installed. If Python is properly installed and added to the PATH, you will see something like:

   ```sh
   C:\Users\YourUsername\AppData\Local\Programs\Python\Python39\python.exe
   ```

### If Python is Not Installed

If the commands above do not return a version number or a path, Python is likely not installed on your system. Here’s how you can install it:

1. **Download Python**:
   - Go to the [official Python website](https://www.python.org/downloads/).
   - Download the latest version of Python.

2. **Run the Installer**:
   - Double-click the downloaded installer file.
   - In the installer window, make sure to check the box that says "Add Python to PATH".
   - Click on "Install Now".

3. **Verify Installation**:
   - After installation, open Command Prompt or PowerShell again.
   - Run `python --version` to verify that Python is installed and properly added to the PATH.

### Troubleshooting

If you encounter any issues, here are some additional tips:

- **Environment Variables**: Ensure that Python's installation directory is added to the `PATH` environment variable. You can manually add it by:
  1. Right-click on `This PC` or `Computer` on the desktop or in File Explorer.
  2. Click `Properties`.
  3. Click `Advanced system settings`.
  4. Click the `Environment Variables` button.
  5. In the `System variables` section, find the `Path` variable, select it, and click `Edit`.
  6. Add the path to the Python executable (e.g., `C:\Users\YourUsername\AppData\Local\Programs\Python\Python39\`).

- **Restart Command Prompt**: If you just installed Python, you may need to restart Command Prompt or PowerShell for the changes to take effect.

By following these steps, you can verify if Python is installed on your Windows system and ensure it's properly configured.

### Downloading the script

Follow these steps to get all python related apps installed and configured.

1. Navigate to where you want the project to live on your system.

```sh
# Find a home for the script
cd /where/you/keep/code
```

2. Clone the project to your computer.

```sh
# Clone the project to your computer
git clone https://github.com/tekNorah/touchpoint-meja.git
```

3. Enter the project's main directory

```sh
# Enter the project folder (where you cloned it)
cd touchpoint-meja
```

### Installing the Packages

On Windows, you can install Python packages using either pip or pipenv. Here's how to do it using pip:

#### Option 1: Using pip

1. **Open Command Prompt**:
   - Press `Win + R`, type `cmd`, and press `Enter`.

2. **Navigate to Your Script's Directory**:
   - Use the `cd` command to change the directory to where your script is located. For example:
     ```sh
     cd C:\path\to\your\script
     ```

3. **Install Required Packages**:
   - Use pip to install the required packages (`pyodbc`, `pandas`, `openpyxl`):
     ```sh
     pip install pyodbc pandas openpyxl
     ```

4. **Run Your Script**:
   - Once the packages are installed, you can run your script:
     ```sh
     python azure_sql_to_excel.py
     ```

#### Option 2: Using pipenv

1. **Open Command Prompt**:
   - Press `Win + R`, type `cmd`, and press `Enter`.

2. **Navigate to Your Script's Directory**:
   - Use the `cd` command to change the directory to where your script is located. For example:
     ```sh
     cd C:\path\to\your\script
     ```

3. **Create a Virtual Environment**:
   - Create a virtual environment using pipenv:
     ```sh
     pipenv install
     ```

4. **Install Required Packages**:
   - Install the required packages (`pyodbc`, `pandas`, `openpyxl`) within the virtual environment:
     ```sh
     pipenv install pyodbc pandas openpyxl
     ```

5. **Activate the Virtual Environment**:
   - Activate the virtual environment:
     ```sh
     pipenv shell
     ```

## Running the Script
1. **Run Your Script**:
   - Once the virtual environment is activated, you can run your script:
     ```sh
     python azure_sql_to_excel.py
     ```