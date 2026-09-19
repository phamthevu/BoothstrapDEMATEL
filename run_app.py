import os
import subprocess
import webbrowser
import time

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    app_path = os.path.join(current_dir, "app.py")

    # chạy streamlit bằng command string (ổn định hơn)
    cmd = f"streamlit run \"{app_path}\" --server.headless=true --server.port=8501"

    # mở browser sau 2s
    subprocess.Popen(cmd, shell=True)

    time.sleep(2)
    webbrowser.open("http://localhost:8501")

    # giữ cửa sổ không tắt
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()