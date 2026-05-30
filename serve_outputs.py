from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
import socket
import os
import subprocess
import shutil
import sys
import time
import re
import threading

# Coerce stdout/stderr to UTF-8 to prevent encoding crash in Windows console
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

OUTPUT_DIR = Path(__file__).resolve().parent / "outputs"
PORT = 8000


def ensure_python_package(package_name: str, import_name: str | None = None) -> bool:
    import_name = import_name or package_name
    try:
        __import__(import_name)
        return True
    except ImportError:
        print(f"[Info] Cài đặt gói Python thiếu: {package_name}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            return True
        except subprocess.CalledProcessError as e:
            print(f"[Error] Không cài được {package_name}: {e}")
            return False


def ensure_python_packages(packages: dict[str, str]) -> bool:
    for package_name, import_name in packages.items():
        if not ensure_python_package(package_name, import_name):
            return False
    return True


def ensure_npm_available() -> bool:
    if shutil.which("npm"):
        return True
    print("[Info] npm không tìm thấy trên hệ thống. Vui lòng cài Node.js để hỗ trợ localtunnel.")
    return False


def ensure_node_package(package_name: str) -> bool:
    if shutil.which(package_name) or shutil.which("npx"):
        return True
    if not ensure_npm_available():
        return False
    print(f"[Info] Cài đặt package Node thiếu: {package_name}...")
    try:
        subprocess.check_call(["npm", "install", "-g", package_name])
        return True
    except subprocess.CalledProcessError as e:
        print(f"[Error] Không cài được {package_name} qua npm: {e}")
        return False


def run_pipeline() -> bool:
    script_path = Path(__file__).resolve().parent / "final_project.py"
    if not script_path.exists():
        print(f"[Error] Không tìm thấy {script_path}")
        return False
    print("[Info] Kiểm tra cú pháp final_project.py...")
    try:
        subprocess.check_call([sys.executable, "-m", "py_compile", str(script_path)], cwd=script_path.parent)
    except subprocess.CalledProcessError as e:
        print(f"[Error] PyCompile failed: {e}")
        return False
    print("[Info] Chạy final_project.py để tạo lại HTML...")
    try:
        subprocess.check_call([sys.executable, str(script_path)], cwd=script_path.parent)
    except subprocess.CalledProcessError as e:
        print(f"[Error] Chạy final_project.py thất bại: {e}")
        return False
    return True


def test_local_url(url: str) -> bool:
    if not ensure_python_package("requests"):
        print("[Warning] Không thể cài requests để test URL.")
        return False
    import requests
    print(f"[Info] Kiểm tra lại trang: {url}")
    try:
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            print("[Success] Trang hoạt động với status 200.")
            return True
        print(f"[Warning] Trang trả về status {resp.status_code}.")
        return False
    except Exception as e:
        print(f"[Error] Không truy cập được URL thử nghiệm: {e}")
        return False

def get_local_ip() -> str:
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


def create_public_tunnel(port: int):
    if not ensure_python_package("pyngrok"):
        print("\n[Warning] Không thể cài pyngrok. Bỏ qua ngrok và thử localtunnel.")
        return None

    try:
        from pyngrok import ngrok, conf
    except ImportError:
        print("\n[Info] pyngrok chưa cài hoặc không thể import. Bỏ qua ngrok và thử localtunnel.")
        return None

    auth_token = os.environ.get("NGROK_AUTH_TOKEN")
    if auth_token:
        conf.get_default().auth_token = auth_token
    else:
        print("\n[Info] NGROK_AUTH_TOKEN chưa có, bỏ qua ngrok và thử localtunnel.")
        return None

    try:
        tunnel = ngrok.connect(port, "http")
        print(f"Open globally: {tunnel.public_url}/london_tfl_map.html")
        print("(Ngrok tunnel active until you stop this script.)")
        return tunnel
    except Exception as e:
        print("\n[Warning] Không thể tạo ngrok tunnel:", e)
        return None


def create_localtunnel(port: int):
    if not ensure_npm_available():
        return None

    if not shutil.which("npx") and not shutil.which("localtunnel"):
        if not ensure_node_package("localtunnel"):
            return None

    default_subdomain = os.environ.get("LT_SUBDOMAIN")
    if not default_subdomain:
        import hashlib
        default_subdomain = f"london-tfl-{hashlib.sha1(str(OUTPUT_DIR).encode()).hexdigest()[:8]}"
    print(f"\n[Info] Đang cố gắng mở public URL qua localtunnel với subdomain cố định: {default_subdomain}...")

    if shutil.which("npx"):
        cmd = [
            "npx",
            "--yes",
            "localtunnel",
            "--port",
            str(port),
            "--print-requests",
            "--subdomain",
            default_subdomain,
        ]
    else:
        cmd = [
            "localtunnel",
            "--port",
            str(port),
            "--print-requests",
            "--subdomain",
            default_subdomain,
        ]
    if sys.platform.startswith("win"):
        cmd = ["cmd", "/c"] + cmd

    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    except Exception as e:
        print("[Warning] Không thể chạy localtunnel:", e)
        return None

    url = None
    start_time = time.time()
    url_pattern = re.compile(r"your url is: (https?://[\w\.-]+)", re.IGNORECASE)

    while time.time() - start_time < 20:
        if proc.stdout is None:
            break
        line = proc.stdout.readline()
        if not line:
            if proc.poll() is not None:
                break
            continue
        print(line.strip())
        match = url_pattern.search(line)
        if match:
            url = match.group(1)
            break

    if url:
        print(f"Open globally: {url}/london_tfl_map.html")
        print("(Localtunnel active until you stop this script.)")
        return proc

    print("[Warning] localtunnel không tạo được URL cố định, thử lại với subdomain ngẫu nhiên...")
    proc.terminate()

    cmd = [
        "npx",
        "--yes",
        "localtunnel",
        "--port",
        str(port),
        "--print-requests",
    ]
    if sys.platform.startswith("win"):
        cmd = ["cmd", "/c"] + cmd

    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    except Exception as e:
        print("[Warning] Không thể chạy localtunnel ngẫu nhiên:", e)
        return None

    url = None
    start_time = time.time()
    while time.time() - start_time < 20:
        if proc.stdout is None:
            break
        line = proc.stdout.readline()
        if not line:
            if proc.poll() is not None:
                break
            continue
        print(line.strip())
        match = url_pattern.search(line)
        if match:
            url = match.group(1)
            break

    if url:
        print(f"Open globally: {url}/london_tfl_map.html")
        print("(Localtunnel active until you stop this script.)")
        return proc

    print("[Warning] localtunnel không tạo được URL trong thời gian chờ.")
    proc.terminate()
    return None


if __name__ == "__main__":
    if not OUTPUT_DIR.exists():
        raise FileNotFoundError(f"Output directory not found: {OUTPUT_DIR}")

    required_packages = {
        "pandas": "pandas",
        "numpy": "numpy",
        "scikit-learn": "sklearn",
        "requests": "requests",
        "pyngrok": "pyngrok",
    }
    if not ensure_python_packages(required_packages):
        print("[Error] Không thể cài đủ các gói Python cần thiết. Dừng script.")
        sys.exit(1)

    if not run_pipeline():
        print("[Error] Pipeline không chạy được. Kiểm tra final_project.py.")
        sys.exit(1)

    handler = SimpleHTTPRequestHandler
    server = None
    actual_port = PORT
    for attempt in range(20):
        try:
            server = ThreadingHTTPServer(("0.0.0.0", actual_port), handler)
            break
        except OSError as e:
            print(f"[Warning] Cổng {actual_port} đã bị chiếm dụng. Đang thử cổng {actual_port + 1}...")
            actual_port += 1
            
    if server is None:
        print("[Error] Không tìm thấy cổng trống để khởi chạy HTTP Server.")
        sys.exit(1)
        
    PORT = actual_port

    os.chdir(OUTPUT_DIR)
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    print("Serving static files from:", OUTPUT_DIR)
    print("Open Map Locally: http://127.0.0.1:%d/london_tfl_map.html" % PORT)
    print("Open Map on LAN: http://%s:%d/london_tfl_map.html" % (get_local_ip(), PORT))

    local_url = f"http://127.0.0.1:{PORT}/london_tfl_map.html"
    test_local_url(local_url)

    fixed = os.environ.get("LT_SUBDOMAIN", "london-tfl-map")
    print(f"Trying to create a public URL automatically with fixed localtunnel subdomain: {fixed}. If it fails, install ngrok or localtunnel.")

    tunnel = create_public_tunnel(PORT)
    localtunnel_proc = None
    if tunnel is None:
        localtunnel_proc = create_localtunnel(PORT)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nServer stopped.")
        server.shutdown()
        server.server_close()
        if tunnel:
            from pyngrok import ngrok
            ngrok.disconnect(tunnel.public_url)
        if localtunnel_proc:
            localtunnel_proc.terminate()
            print("Localtunnel stopped.")
