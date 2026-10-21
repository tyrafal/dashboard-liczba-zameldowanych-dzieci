@echo off
chcp 65001 >nul
echo ============================================================================
echo  Aktualizacja zbiorczego raportu Excel i Dashboardu HTML (Roczniki .xlsx)
echo ============================================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0generuj_raport.ps1"
if %ERRORLEVEL% EQU 0 (
    echo.
    echo [SUKCES] Dane zostaly pomyslnie zaktualizowane!
    echo  - Zbiorczy Excel: zbiorczy_raport_2001_2026.xlsx
    echo  - Dashboard HTML: dashboard_zameldowani.html
) else (
    echo.
    echo [BLAD] Wystapil problem podczas aktualizacji danych.
)
echo.
pause
