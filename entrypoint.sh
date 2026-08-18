#!/bin/bash
set -e

exec pwsh -NoLogo -NonInteractive -File /app/listener.ps1
