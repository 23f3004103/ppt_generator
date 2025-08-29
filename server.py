#!/usr/bin/env python3
"""
PowerPoint Generator & LLM Agent Demo Server
Simple HTTP server to serve both applications
"""

import http.server
import socketserver
import webbrowser
import os
import sys
from pathlib import Path

def main():
    print("=" * 60)
    print("   PowerPoint Generator & LLM Agent Demo Server")
    print("=" * 60)
    print()
    
    # Set port
    PORT = 8080
    
    # Change to the script directory
    script_dir = Path(__file__).parent
    os.chdir(script_dir)
    
    # Create server
    Handler = http.server.SimpleHTTPRequestHandler
    
    try:
        with socketserver.TCPServer(("", PORT), Handler) as httpd:
            print(f"🚀 Server starting on http://localhost:{PORT}")
            print()
            print("Available Applications:")
            print(f"  🎨 PowerPoint Generator: http://localhost:{PORT}/app.html")
            print(f"  🤖 LLM Agent Demo:       http://localhost:{PORT}/index.html")
            print(f"  📋 Quick Demo:           http://localhost:{PORT}/demo.html")
            print(f"  📚 Examples Gallery:     http://localhost:{PORT}/examples.html")
            print()
            print("� PowerPoint Generator Features:")
            print("  - Transform text into professional presentations")
            print("  - AI-powered content analysis and slide structuring")
            print("  - Template-based styling with asset preservation")
            print("  - Multi-LLM support (OpenAI, Anthropic, Google)")
            print("  - Real-time preview and download capabilities")
            print()
            print("🛠️ LLM Agent Demo Features:")
            print("  - Multi-tool integration (Search, Calculate, Code, API)")
            print("  - Intelligent task planning and execution")
            print("  - Real-time tool monitoring and execution logs")
            print("  - Interactive chat interface with markdown support")
            print()
            print("💡 Try these PowerPoint Generator examples:")
            print("  • Paste business plan → 'turn into investor pitch deck'")
            print("  • Paste documentation → 'create training presentation'")
            print("  • Paste research paper → 'make it a conference talk'")
            print()
            print("💡 Try these LLM Agent examples:")
            print("  • 'Search for latest AI developments and analyze trends'")
            print("  • 'Calculate compound interest and create visualization'")
            print("  • 'Generate prime numbers and create a chart'")
            print()
            print("Press Ctrl+C to stop the server")
            print("-" * 60)
            
            # Try to open browser automatically
            try:
                print("🌐 Opening PowerPoint Generator in your default browser...")
                webbrowser.open(f'http://localhost:{PORT}/app.html')
            except:
                print("Could not open browser automatically")
                print(f"Please navigate to: http://localhost:{PORT}/app.html")
            
            print()
            print("Server is running... (Ctrl+C to stop)")
            httpd.serve_forever()
            
    except KeyboardInterrupt:
        print("\n\n🛑 Server stopped by user")
        print("Thanks for trying the PowerPoint Generator and LLM Agent demos!")
    except OSError as e:
        if e.errno == 48:  # Address already in use
            print(f"❌ Port {PORT} is already in use!")
            print("Try stopping other servers or use a different port:")
            print(f"   python server.py --port 8081")
        else:
            print(f"❌ Error starting server: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
