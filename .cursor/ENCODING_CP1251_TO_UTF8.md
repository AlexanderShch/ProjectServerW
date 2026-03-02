# Перекодировка файлов ProjectServerW из CP1251 в UTF-8

Проверенный способ (не портит русские буквы):

1. Создать скрипт `to_utf8.ps1` в корне проекта:

```powershell
$path = Join-Path (Split-Path -Parent $PSCommandPath) "DataForm.h"
$cp1251 = [System.Text.Encoding]::GetEncoding(1251)
$utf8Bom = [System.Text.UTF8Encoding]::new($true)
$content = [System.IO.File]::ReadAllText($path, $cp1251)
[System.IO.File]::WriteAllText($path, $content, $utf8Bom)
Write-Host "Done"
```

2. Запустить:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "D:\...\ProjectServerW\to_utf8.ps1"
```

3. Удалить скрипт после выполнения.

Важно: весь файл должен быть в одной кодировке (CP1251). При смешанной кодировке конец файла может стать нечитаемым.
