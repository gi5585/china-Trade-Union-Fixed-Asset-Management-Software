@echo off
chcp 65001 >nul
title 工会固定资产管理系统 V11.0
cd /d "%~dp0"
start pythonw.exe main.py
exit
