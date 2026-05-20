@echo off
title Sincronizador de Base de Datos Excel SBS + Gemini AI
color 0b
echo ==============================================================================
echo       Sincronizador de Base de Datos Excel SBS (Superintendencia de Peru)
echo ==============================================================================
echo.
echo Iniciando proceso de extraccion incremental de nuevos dias...
echo.
python sbs_excel_extractor.py
echo.
echo ==============================================================================
echo Proceso finalizado.
echo ==============================================================================
pause
