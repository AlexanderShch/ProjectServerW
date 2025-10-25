@echo off
echo Clearing IntelliSense cache...

REM �������� ����� .vs (�������� ��� IntelliSense)
if exist ".vs" (
    echo Deleting .vs folder...
    rmdir /s /q ".vs"
    echo .vs folder deleted.
) else (
    echo .vs folder not found.
)

REM �������� ������ .ipch (IntelliSense Precompiled Headers)
if exist "*.ipch" (
    echo Deleting .ipch files...
    del /q "*.ipch"
    echo .ipch files deleted.
)

REM �������� ������ .db (���� ������ IntelliSense)
if exist "*.db" (
    echo Deleting .db files...
    del /q "*.db"
    echo .db files deleted.
)

echo.
echo IntelliSense cache cleared!
echo Please reopen Visual Studio and rebuild the solution.
echo.
pause

