@echo.                                                                           >> xxxdevs.log
@echo.                                                                           >> xxxdevs.log
@echo "Updated on %USERDOMAIN%\%COMPUTERNAME% by %USERNAME% on %DATE% at %TIME%" >> xxxdevs.log
@echo.                                                                           >> xxxdevs.log
@echo PROCESSOR_ARCHITECTURE %PROCESSOR_ARCHITECTURE%                            >> xxxdevs.log
@echo PROCESSOR_ARCHITEW3264 %PROCESSOR_ARCHITEW3264%                            >> xxxdevs.log
@echo PROCESSOR_IDENTIFIER   %PROCESSOR_IDENTIFIER%                              >> xxxdevs.log
@echo CLIENTNAME             %CLIENTNAME%                                        >> xxxdevs.log
@echo SESSIONNAME            %SESSIONNAME%                                       >> xxxdevs.log
@echo NUMBER_OF_PROCESSORS   %NUMBER_OF_PROCESSORS%                              >> xxxdevs.log
@echo.                                                                           >> xxxdevs.log
@setlocal
@set siv_inf_alt=%SystemRoot%\ServicePackFiles
@if not exist %siv_inf_alt% set siv_inf_alt=
@if exist mondevs.exe mondevs %SystemRoot%\inf %siv_inf_alt% >> xxxdevs.log
@if exist mondevs.exe echo.                                  >> xxxdevs.log
@if exist pcidevs.exe pcidevs %SystemRoot%\inf %siv_inf_alt% >> xxxdevs.log
@if exist pcidevs.exe echo.                                  >> xxxdevs.log
@if exist pcmdevs.exe pcmdevs %SystemRoot%\inf %siv_inf_alt% >> xxxdevs.log
@if exist pcmdevs.exe echo.                                  >> xxxdevs.log
@if exist pnpdevs.exe pnpdevs %SystemRoot%\inf %siv_inf_alt% >> xxxdevs.log
@if exist pnpdevs.exe echo.                                  >> xxxdevs.log
@if exist usbdevs.exe usbdevs %SystemRoot%\inf %siv_inf_alt% >> xxxdevs.log
@if exist usbdevs.exe echo.                                  >> xxxdevs.log
@endlocal
@setlocal
@if "%USERDOMAIN%"=="%COMPUTERNAME%"    set siv_domain=
@                                       set siv_image=siv32l
@if exist %windir%\system32\devmgmt.msc set siv_image=siv32x
@if "%PROCESSOR_ARCHITECTURE%"=="ALPHA" set siv_image=siv32a
@if "%PROCESSOR_ARCHITECTURE%"=="AMD64" set siv_image=siv64x
@if "%PROCESSOR_ARCHITEW3264%"=="AMD64" set siv_image=siv64x
@if "%PROCESSOR_ARCHITECTURE%"=="IA64"  set siv_image=siv64i
@if "%PROCESSOR_ARCHITEW3264%"=="IA64"  set siv_image=siv64i
@if not exist %siv_image%.exe           set siv_image=siv32x
@if     exist %siv_image%.exe echo %siv_image% -save        >> xxxdevs.log
@if     exist %siv_image%.exe      %siv_image% -save | more >> xxxdevs.log
@endlocal