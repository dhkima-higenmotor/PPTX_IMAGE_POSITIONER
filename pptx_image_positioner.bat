@echo off
call ./.venv/Scripts/activate
call uv run pptx_image_positioner.py
REM pause