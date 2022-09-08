@echo off

"%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe" -noexit -executionpolicy bypass -command "import-module '%~dp0Warranty.ps1'"