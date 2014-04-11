@echo.                                                                          > SIVDBG.log
@echo.                                                                         >> SIVDBG.log
@echo Testrun on %USERDOMAIN%\%COMPUTERNAME% by %USERNAME% on %DATE% at %TIME% >> SIVDBG.log
@echo.                                                                         >> SIVDBG.log
@echo PROCESSOR_ARCHITECTURE %PROCESSOR_ARCHITECTURE%                          >> SIVDBG.log
@echo PROCESSOR_ARCHITEW3264 %PROCESSOR_ARCHITEW3264%                          >> SIVDBG.log
@echo PROCESSOR_IDENTIFIER   %PROCESSOR_IDENTIFIER%                            >> SIVDBG.log
@echo CLIENTNAME             %CLIENTNAME%                                      >> SIVDBG.log
@echo SESSIONNAME            %SESSIONNAME%                                     >> SIVDBG.log
@echo NUMBER_OF_PROCESSORS   %NUMBER_OF_PROCESSORS%                            >> SIVDBG.log
@echo.                                                                         >> SIVDBG.log
@setlocal
@                                       set siv_exe=siv32l
@if exist %windir%\system32\devmgmt.msc set siv_exe=siv32x
@                                       set siv_nat=%siv_exe%
@if "%PROCESSOR_ARCHITECTURE%"=="ALPHA" set siv_nat=siv32a
@if "%PROCESSOR_ARCHITECTURE%"=="AMD64" set siv_nat=siv64x
@if "%PROCESSOR_ARCHITEW3264%"=="AMD64" set siv_nat=siv64x
@if "%PROCESSOR_ARCHITECTURE%"=="IA64"  set siv_nat=siv64i
@if "%PROCESSOR_ARCHITEW3264%"=="IA64"  set siv_nat=siv64i
@if exist %siv_nat%.exe                 set siv_exe=%siv_nat%
@if exist %siv_exe%.exe (
%siv_exe% -DBGINI -DBGHAL -NOACPI -DBGSDM -NOSMART -DBGSMB -NOSMBUS -NOSENSORS >> SIVDBG.log
%siv_exe% -DBGINI -DBGHAL -NOACPI -DBGSDM -NOSMART -DBGSMB -NOSMBUS            >> SIVDBG.log
%siv_exe% -DBGINI -DBGHAL -NOACPI -DBGSDM -NOSMART -DBGSMB                     >> SIVDBG.log
%siv_exe% -DBGINI -DBGHAL -NOACPI -DBGSDM                                      >> SIVDBG.log
%siv_exe% -DBGINI -DBGHAL                                                      >> SIVDBG.log
) else (
@echo.
@echo Failed to find %siv_exe%.exe on %USERDOMAIN%\%COMPUTERNAME% by %USERNAME% on %DATE% at %TIME%
)
@endlocal