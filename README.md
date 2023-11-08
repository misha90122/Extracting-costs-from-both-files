Step 1: Download and Install Python:

Visit the official Python website at https://www.python.org/.
In the "Downloads" section, select the latest stable version of Python for your operating system (Windows, macOS, or Linux).
Download the executable file for Python installation.
After downloading, run the executable file and follow the Python installer's instructions.
Install Python with the default settings. You may also be prompted to add Python to your system's PATH variable, providing convenient access to Python from any folder in the command line (terminal).
Step 2: Check Python Installation:

Open the command prompt (for Windows) or terminal (for macOS and Linux).
Enter the command python --version or python3 --version to check the installed Python version. If you see a version number (e.g., Python 3.9.6), it means Python is installed successfully.
Step 3: Editing the Script and Preparing Files for Execution:

Save the script "find_expenses_for_numbers.py" in a separate folder along with two Excel files ("диф звіт травень.xlsx" and "диф рахунок червень.xlsx") and the file "kyivstar_numbers.txt".
Open the "kyivstar_numbers.txt" file and input phone numbers, each on a separate line without extra characters.
Replace the file paths in the "find_expenses_for_numbers.py" script with your own paths:
input_file_path1: Path to the first Excel file ("диф звіт травень.xlsx").
input_file_path2: Path to the second Excel file ("диф рахунок червень.xlsx").
output_file_path: Path to the output Excel file (can be "output.xlsx" or any other filename).
numbers_file_path: Path to the "kyivstar_numbers.txt" file.
Step 4: Running the Script:

Open the command prompt (for Windows) or terminal (for macOS and Linux).
Navigate to the folder where the script and Excel files, as well as "kyivstar_numbers.txt," are saved. Use the cd <folder path> command for this.
Run the script by entering the command: python find_expenses_for_numbers.py (or python3 find_expenses_for_numbers.py for Linux and macOS).
The script will process the data, find expenses for the first and second month for each phone number, and save the results in the output Excel file "output.xlsx."
Now that you have completed these steps, you can successfully use this script to find expenses based on the provided Excel files and the phone numbers file. This documentation is formatted for GitHub.
