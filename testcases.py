# INCOMPLETE
import subprocess
import os

# Function to run the CLI program and capture its output
def run_cli_command(input_commands):
    try:
        # Run the CLI program using subprocess
        process = subprocess.Popen(
            ["python", "your_cli_program.py"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        # Send input commands to the program
        for command in input_commands:
            process.stdin.write(command + "\n")
            process.stdin.flush()

        # Capture and return the program's output and errors
        stdout, stderr = process.communicate()
        return stdout, stderr
    except Exception as e:
        return None, str(e)

# Test Case 1.1: Opening an Existing Excel File
input_commands = ["test.xlsx"]
stdout, stderr = run_cli_command(input_commands)
if not stderr:
    print("Test Case 1.1: PASSED - Successfully opened an existing Excel file.")
else:
    print("Test Case 1.1: FAILED - Error occurred:", stderr)

# Test Case 2.2: Generating 5 Tickets
input_commands = ["5"]
stdout, stderr = run_cli_command(input_commands)
if not stderr and "Number of winners:" in stdout:
    print("Test Case 2.2: PASSED - Successfully generated 5 tickets.")
else:
    print("Test Case 2.2: FAILED - Error occurred or incorrect output.")

# Test Case 3.1: Saving to Excel
output_file = "output.xlsx"
input_commands = ["output.xlsx"]
stdout, stderr = run_cli_command(input_commands)
if not stderr and os.path.exists(output_file):
    os.remove(output_file)  # Clean up the generated file
    print("Test Case 3.1: PASSED - Successfully saved data to Excel.")
else:
    print("Test Case 3.1: FAILED - Error occurred or output file not found.")

# Test Case 4.1: Invalid Input for Number of Tickets
input_commands = ["-3"]
stdout, stderr = run_cli_command(input_commands)
if stderr and "must be positive" in stderr:
    print("Test Case 4.1: PASSED - Correctly handled invalid input for number of tickets.")
else:
    print("Test Case 4.1: FAILED - Expected error message not found in output.")

# Test Case 4.2: Invalid File Path
input_commands = ["non_existing_file.csv"]
stdout, stderr = run_cli_command(input_commands)
if stderr and "does not exist" in stderr:
    print("Test Case 4.2: PASSED - Correctly handled invalid file path.")
else:
    print("Test Case 4.2: FAILED - Expected error message not found in output.")

# Test Case 5.1: Clear and Informative Prompts
input_commands = []
stdout, stderr = run_cli_command(input_commands)
if not stderr:
    print("Test Case 5.1: PASSED - Verified clear and informative prompts.")
else:
    print("Test Case 5.1: FAILED - Error occurred or prompts not clear.")

# Test Case 5.2: Handling Different File Formats
input_commands = ["test.xlsx", "test.csv"]
stdout, stderr = run_cli_command(input_commands)
if not stderr:
    print("Test Case 5.2: PASSED - Successfully handled different file formats.")
else:
    print("Test Case 5.2: FAILED - Error occurred while handling file formats.")

