:��������� ������ nvidia � ������ ����� �������
sc config NVSvc start= demand

:������� �� ����������� ������� ������
REG DELETE "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /v "NvCplDaemon" /f

:������� ����� ������� ��� ���������� ��������� �������� NVIDIA
REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Reinstall"

sleep 5