import subprocess
import sys

required = [
    "xlwings",
    "pyinstaller"
]


def install(package):
    print(f"📦 Installing {package}...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


def main():
    print("🔍 Checking and installing required packages...\n")
    for package in required:
        try:
            __import__(package)
            print(f"✅ {package} is already installed.")
        except ImportError:
            install(package)

    print("\n🎉 All required packages are installed and ready to use!")


if __name__ == "__main__":
    main()
