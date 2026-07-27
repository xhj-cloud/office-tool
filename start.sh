#!/bin/bash
# Office Tools MCP Server launcher
# 用于 CherryStudio MCP 配置中的启动命令

cd "$(dirname "$0")"
exec venv/bin/python server.py "$@"
