@echo off
cd /d C:\Users\dasilva.lucas\Documents\MCP\Office-Word-MCP-Server
call venv\Scripts\activate.bat
set PYTHONPATH=C:\Users\dasilva.lucas\Documents\MCP\Office-Word-MCP-Server
python word_document_server\main.py
