import subprocess

def run_command_in_cmd():
    command = ["python", "-u", "main.py"]
    subprocess.Popen(command, shell=True)

if __name__ == "__main__":
    run_command_in_cmd()
