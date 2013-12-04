:Переводим службу nvidia в ручной режим запуска
sc config NVSvc start= demand

:Удаляем из автозапуска сбойную запись
REG DELETE "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /v "NvCplDaemon" /f

:Создаем ветку реестра для корректной установки драйвера NVIDIA
REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Reinstall"

sleep 5