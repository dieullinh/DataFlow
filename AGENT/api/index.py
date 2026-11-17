# -*- coding: utf-8 -*-
"""
Vercel Serverless Function Entry Point for DataFlow Platform
Compatible with standard Vercel Python runtime and BaseHTTPRequestHandler.
"""
from http.server import BaseHTTPRequestHandler
import json

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        response = {
            "status": "healthy",
            "app": "DataFlow",
            "version": "2.0.0",
            "message": "DataFlow Analytics API is active."
        }
        self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))

    def do_POST(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        response = {
            "status": "success",
            "message": "Data received successfully."
        }
        self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))
