@echo off
reg.exe add "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /f /ve
cmd /K echo 実行完了：右クリックメニューを旧仕様（Windows10）に変更しました。
