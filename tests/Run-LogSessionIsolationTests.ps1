$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$vs = & $vswhere -latest -requires Microsoft.Component.MSBuild -property installationPath
if (-not $vs) { throw 'Visual Studio with MSBuild is required.' }
$msbuild = Join-Path $vs 'MSBuild\Current\Bin\MSBuild.exe'
$testOutput = Join-Path ([IO.Path]::GetTempPath()) ('ProjectServerW-LogTests-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testOutput | Out-Null

Push-Location $repo
try {
    & $msbuild ProjectServerW.sln /p:Configuration=Debug /p:Platform=x64 /m /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) { throw 'Debug x64 build failed.' }

    # Mixed C++/CLI executables cannot be loaded with Assembly.LoadFrom. Link the
    # exact application objects/resources as a DLL for the reflection-based harness.
    Import-Module (Join-Path $vs 'Common7\Tools\Microsoft.VisualStudio.DevShell.dll')
    Enter-VsDevShell -VsInstallPath $vs -SkipAutomaticLocation -DevCmdArguments '-arch=x64 -host_arch=x64'
    $sendTests = Join-Path $testOutput 'SocketSendTests.exe'
    & cl.exe /nologo /EHsc /std:c++14 (Join-Path $PSScriptRoot 'SocketSendTests.cpp') "/Fo$testOutput\SocketSendTests.obj" "/Fe$sendTests" /link ws2_32.lib
    if ($LASTEXITCODE -ne 0) { throw 'Socket send test compilation failed.' }
    & $sendTests
    if ($LASTEXITCODE -ne 0) { throw 'Socket send tests failed.' }
    $objectDirectory = Join-Path $repo 'x64\Debug\ProjectServerW'
    $testObjects = @(Get-ChildItem -LiteralPath $objectDirectory -Filter '*.obj' | ForEach-Object FullName)
    $subject = Join-Path $testOutput 'LogSessionTestSubject.dll'
    $resources = @(
        '/ASSEMBLYRESOURCE:' + (Join-Path $objectDirectory 'ProjectServerW.DataForm.resources')
        '/ASSEMBLYRESOURCE:' + (Join-Path $objectDirectory 'ProjectServerW.MyForm.resources')
    )
    & link.exe /DLL /NOLOGO /MACHINE:X64 /INCREMENTAL:NO "/OUT:$subject" $resources $testObjects ws2_32.lib Ole32.lib
    if ($LASTEXITCODE -ne 0) { throw 'Test subject link failed.' }

    foreach ($dependency in @('Microsoft.Office.Interop.Excel.dll', 'stdole.dll')) {
        Copy-Item -LiteralPath (Join-Path $objectDirectory $dependency) -Destination (Join-Path $testOutput $dependency)
    }
    $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
    $harness = Join-Path $testOutput 'LogSessionIsolationTests.exe'
    & $compiler /nologo /platform:x64 /r:System.Data.dll /r:System.Windows.Forms.dll "/out:$harness" (Join-Path $PSScriptRoot 'LogSessionIsolationTests.cs')
    if ($LASTEXITCODE -ne 0) { throw 'Test harness compilation failed.' }
    & $harness $subject
    if ($LASTEXITCODE -ne 0) { throw 'Log session isolation tests failed.' }
}
finally { Pop-Location }
