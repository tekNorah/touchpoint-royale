# `touchpoint-meja`

![Static Badge](https://img.shields.io/badge/goal-Transform_Azure_SQL_Data_to_Excel-purple)
<br />
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

<p class="align center">
<h4><code>touchpoint-meja</code> is a Python script for transforming touchpoint data for the <code>NCR-MothersEntJA</code> database to an Excel Workbook.</h4>
</p>

[What and Why](#whatandwhy) •
[Quickstart](#quickstart) •
[Structure](#structure) •
[Meta](#meta)

</div>

## Navigation

- [Introduction Videos](#introduction-videos)
- [What and Why](#what-and-why)
- [Quickstart](#quickstart)
  - [Setting up the fabric commands](#setting-up-the-fabric-commands)
  - [Using the fabric client](#using-the-fabric-client)
  - [Just use the Patterns](#just-use-the-patterns)
  - [Create your own Fabric Mill](#create-your-own-fabric-mill)
- [Structure](#structure)
  - [Components](#components)
  - [CLI-native](#cli-native)
  - [Directly calling Patterns](#directly-calling-patterns)
- [Meta](#meta)
  - [Primary contributors](#primary-contributors)

<br />

## What and why

The purpose of this Python scrip is to 

## Quickstart

The most feature-rich way to use Fabric is to use the `fabric` client, which can be found under <a href="https://github.com/danielmiessler/fabric/tree/main/installer/client">`/client`</a> directory in this repository.

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

```bash
# Find a home for the script
cd /where/you/keep/code
```

2. Clone the project to your computer.

```bash
# Clone the project to your computer
git clone https://github.com/tekNorah/touchpoint-meja.git
```

3. Enter the project's main directory

```bash
# Enter the project folder (where you cloned it)
cd touchpoint-meja
```

