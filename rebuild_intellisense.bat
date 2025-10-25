@echo off
echo Clearing IntelliSense cache...

REM Удаление папки .vs (содержит кэш IntelliSense)
if exist ".vs" (
    echo Deleting .vs folder...
    rmdir /s /q ".vs"
    echo .vs folder deleted.
) else (
    echo .vs folder not found.
)

REM Удаление файлов .ipch (IntelliSense Precompiled Headers)
if exist "*.ipch" (
    echo Deleting .ipch files...
    del /q "*.ipch"
    echo .ipch files deleted.
)

REM Удаление файлов .db (база данных IntelliSense)
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

