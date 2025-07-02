import subprocess
import sys

def install_requirements():
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("All dependencies are installed.")
    except subprocess.CalledProcessError:
        print("Error installing dependencies.")

if __name__ == "__main__":
    install_requirements()