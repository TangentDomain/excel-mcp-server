#!/bin/bash
cd /root/.openclaw/workspace/excel-mcp-server
python3 test_fixes.py 2>&1 | tee test_fixes_output.log
