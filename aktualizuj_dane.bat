@echo off
chcp 65001 >nul
echo ============================================================================
echo  Aktualizacja zbiorczych raportow Excel i Dashboardow HTML (Roczniki .xlsx)
echo ============================================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0generuj_raport.ps1"
if %ERRORLEVEL% EQU 0 (
    echo.
    echo [SUKCES] Dane zostaly pomyslnie zaktualizowane dla obu wersji!
    echo  1. Wersja 2001-2026:
    echo     - Dashboard HTML: dashboard_zameldowani_2001_2026.html
    echo     - Zbiorczy Excel: zbiorczy_raport_2001_2026.xlsx
    echo  2. Wersja 2021-2026:
    echo     - Dashboard HTML: dashboard_zameldowani_2021_2026.html
    echo     - Zbiorczy Excel: zbiorczy_raport_2021_2026.xlsx
) else (
    echo.
    echo [BLAD] Wystapil problem podczas aktualizacji danych.
)
echo.
pause
