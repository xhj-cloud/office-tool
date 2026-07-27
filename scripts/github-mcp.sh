#!/bin/bash
export GITHUB_TOKEN=$(/opt/homebrew/bin/gh auth token)
exec /opt/homebrew/bin/npx -y @modelcontextprotocol/server-github
