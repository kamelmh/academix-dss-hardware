import subprocess
import sys

# Start Streamlit
cmd = [
    sys.executable, "-m", "streamlit", "run",
    "C:/Users/Admin/projects/active/academix-dss-hardware/streamlit-dashboard/app.py",
    "--server.headless", "true",
    "--server.port", "8501",
    "--server.address", "localhost"
]

print("Starting Streamlit...")
print(f"URL: http://localhost:8501")
print("Press Ctrl+C to stop")

subprocess.run(cmd)
