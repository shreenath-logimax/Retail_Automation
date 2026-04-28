# Setup Workflow for Retail Automation

Follow these instructions to set up the **Retail Automation** application on a new system after cloning the repository.

## Prerequisites
1. **Python 3.8+**: Ensure Python is installed and added to your system's PATH. You can download it from [python.org](https://www.python.org/).
2. **Git**: Ensure Git is installed to clone the repository.
3. **Chrome Browser**: The application relies on Selenium, which requires a browser. Chrome is recommended.

## Installation Steps

### Option 1: Automated Setup (Recommended for Windows)
We have provided a batch script that will automatically create a Python virtual environment and install all necessary dependencies.

1. Open the cloned folder `Retail_Automation` in your File Explorer.
2. Double-click the `setup.bat` file.
3. Wait for the script to finish. It will create a `.venv` folder and install everything listed in `requirements.txt`.

### Option 2: Manual Setup
If you prefer to set up the environment manually via the Command Prompt or PowerShell, follow these steps:

1. **Open your terminal** and navigate to the project directory:
   ```cmd
   cd path\to\Retail_Automation
   ```

2. **Create a Virtual Environment**:
   ```cmd
   python -m venv .venv
   ```

3. **Activate the Virtual Environment**:
   - On Windows (Command Prompt):
     ```cmd
     .venv\Scripts\activate.bat
     ```
   - On Windows (PowerShell):
     ```powershell
     .venv\Scripts\Activate.ps1
     ```

4. **Install Dependencies**:
   ```cmd
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```

## Running the Application
Whenever you open a new terminal to run the application, you **must activate the virtual environment** first:
```cmd
.venv\Scripts\activate
```

Then you can execute your scripts (e.g., `cd Sparqla` and `python main.py` or whatever your entry point is).
