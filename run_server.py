import uvicorn
import socket

def get_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "10.0.7.26"

if __name__ == "__main__":
    local_ip = get_ip()
    print("=" * 65)
    print("  *** VISTEX AGREEMENT FORMAT CONVERTER WEB SERVER ***")
    print("=" * 65)
    print("  Status:        ONLINE")
    print("  Host:        0.0.0.0:5000")
    print(f"  Local Access:  http://localhost:5000")
    print(f"  Network IP:    http://{local_ip}:5000")
    print(f"  Internal IP:   http://10.0.7.26:5000")
    print("-" * 65)
    print("  Default Login Credentials:")
    print("  Username: admin")
    print("  Password: vistex2026")
    print("=" * 65)
    print()

    uvicorn.run("app:app", host="0.0.0.0", port=5000, log_level="info")
