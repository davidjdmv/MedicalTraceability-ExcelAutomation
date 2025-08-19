@echo off
setlocal EnableDelayedExpansion
if exist .env (
  for /f "usebackq tokens=* delims=" %%a in (".env") do set %%a
)
python -m src.app
