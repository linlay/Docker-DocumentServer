#!/usr/bin/env bash
set -euo pipefail

copilot_environment_file=/run/onlyoffice-copilot.env
while IFS= read -r -d '' assignment; do
  export "$assignment"
done < "$copilot_environment_file"

exec /usr/bin/python3 /opt/onlyoffice-copilot/copilot_server.py
