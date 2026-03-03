# Перекодировка файлов ProjectServerW из CP1251 в UTF-8

## Почему скрипт портил файл

Если файл **уже в UTF-8** (с BOM или без), а скрипт читает его **как CP1251**, то байты UTF-8 интерпретируются как однобайтовые символы → получается мусор («Р В РЎвЂ» и т.п.), он сохраняется и файл портится. Поэтому перед перекодировкой нужно **проверять кодировку**: при наличии BOM UTF-8 не перекодировать.

## Безопасный скрипт (с проверкой BOM)

Создать `to_utf8.ps1` в корне проекта:

```powershell
$fileName = "DataForm.cpp"
$dir = Split-Path -Parent $PSCommandPath
$path = Join-Path $dir $fileName

$bytes = [System.IO.File]::ReadAllBytes($path)
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    Write-Host "File already has UTF-8 BOM, skip."
    exit 0
}

$cp1251 = [System.Text.Encoding]::GetEncoding(1251)
$utf8Bom = [System.Text.UTF8Encoding]::new($true)
$content = $cp1251.GetString($bytes)
[System.IO.File]::WriteAllText($path, $content, $utf8Bom)
Write-Host "Done"
```

Запуск:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "D:\...\ProjectServerW\to_utf8.ps1"
```

Важно: перекодировать только файл, который **сейчас в CP1251** (без BOM в начале). Если откатили из git и получили обратно UTF-8 — скрипт без проверки BOM снова испортит файл.
