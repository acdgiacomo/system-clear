@echo off
PowerShell -Command "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0system-clear.ps1\"' -Verb RunAs"