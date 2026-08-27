import http.server
import socketserver
import webbrowser
import os

PORT = 8080
DIRECTORY = r"C:\Users\927632_st.tc\.gemini\antigravity\scratch\gy61-vibration-analysis"

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

if __name__ == "__main__":
    os.chdir(DIRECTORY)
    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        url = f"http://localhost:{PORT}/index.html"
        print(f"GY-61 Analysis Web App Server started at: {url}")
        webbrowser.open(url)
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")
