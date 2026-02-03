$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Resolve-Path (Join-Path $scriptDir "..")

Set-Location $repoRoot

$appName = "cotacoes_moedas"
$appPublisher = "davidsc"
$appSupportUrl = "mailto:davidsc@zanotti.com.br"
$iconPath = Join-Path $repoRoot "imagem_ico\\finance.ico"
$innoSetupPath = $env:INNO_SETUP_PATH
$signingPfxPath = $env:SIGNING_PFX_PATH
$signingPassword = $env:SIGNING_PFX_PASSWORD
$signingTimestampUrl = $env:SIGNING_TIMESTAMP_URL
$entryPoint = "main.py"
if (-not $signingTimestampUrl) {
    $signingTimestampUrl = "http://timestamp.digicert.com"
}

$appVersion = @'
import tomllib
from pathlib import Path

data = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))
print(data["project"]["version"])
'@ | poetry run python -
$appVersion = $appVersion.Trim()

$fileVersion = $appVersion
if ($fileVersion -match '^\d+(\.\d+){2}$') {
    $fileVersion = "$fileVersion.0"
} elseif ($fileVersion -notmatch '^\d+(\.\d+){3}$') {
    $fileVersion = "0.0.0.0"
}

$distFolderName = "{0}.dist" -f ([IO.Path]::GetFileNameWithoutExtension($entryPoint))
$distDir = Join-Path $repoRoot "dist\\$distFolderName"

function Get-SigningPassword {
    param ([string]$PasswordValue)

    if ($PasswordValue) {
        return $PasswordValue
    }
    $secure = Read-Host -AsSecureString "Senha do PFX (assinatura)"
    if (-not $secure) {
        return $null
    }
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}

function Invoke-SignFiles {
    param (
        [string[]]$Paths,
        [string]$PfxPath,
        [string]$Password,
        [string]$TimestampUrl,
        [System.Management.Automation.CommandInfo]$SignTool
    )

    foreach ($path in $Paths) {
        if (-not (Test-Path $path)) {
            continue
        }
        & $SignTool.Source sign `
            /f $PfxPath `
            /p $Password `
            /tr $TimestampUrl `
            /td sha256 `
            /fd sha256 `
            $path
    }
}

function Invoke-RobocopyMirror {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Source,
        [Parameter(Mandatory = $true)]
        [string]$Destination
    )

    if (-not (Test-Path $Source)) {
        Write-Host "Robocopy ignorado: fonte nao encontrada em: $Source"
        return
    }

    if (-not (Test-Path $Destination)) {
        New-Item -ItemType Directory -Path $Destination | Out-Null
    }

    & robocopy $Source $Destination /MIR
    $exitCode = $LASTEXITCODE

    if ($exitCode -ge 8) {
        throw "Robocopy falhou com exit code $exitCode (Source: $Source, Destination: $Destination)"
    }
}

poetry install

$env:PLAYWRIGHT_BROWSERS_PATH = (Join-Path $repoRoot "ms-playwright")
poetry run playwright install chromium

$env:PLAYWRIGHT_BROWSERS_PATH = (Join-Path $repoRoot "ms-playwright")
$nuitkaArgs = @(
    "--standalone",
    "--output-dir=dist",
    "--output-filename=cotacoes-moedas",
    "--include-data-dir=planilhas=planilhas",
    "--include-data-dir=ms-playwright=ms-playwright",
    "--windows-company-name=$appPublisher",
    "--windows-product-name=$appName",
    "--windows-file-description=$appName",
    "--windows-product-version=$fileVersion",
    "--windows-file-version=$fileVersion"
)
if (Test-Path $iconPath) {
    $nuitkaArgs += "--windows-icon-from-ico=$iconPath"
}
poetry run python -m nuitka @nuitkaArgs $entryPoint

Write-Host "Sincronizando ms-playwright/ e planilhas/ no dist (robocopy)..."
Invoke-RobocopyMirror -Source (Join-Path $repoRoot "ms-playwright") -Destination (Join-Path $distDir "ms-playwright")
Invoke-RobocopyMirror -Source (Join-Path $repoRoot "planilhas") -Destination (Join-Path $distDir "planilhas")

$portableZipPath = Join-Path $repoRoot "dist\\cotacoes-moedas-portable.zip"
if (Test-Path $portableZipPath) {
    Remove-Item -Force $portableZipPath
}
Compress-Archive -Path (Join-Path $distDir "*") -DestinationPath $portableZipPath -CompressionLevel Optimal
Write-Host "Portable ZIP gerado em: $portableZipPath"

$signingPfxResolved = $null
if ($signingPfxPath) {
    $signingPfxResolved = Resolve-Path $signingPfxPath -ErrorAction SilentlyContinue
    if (-not $signingPfxResolved) {
        Write-Host "PFX nao encontrado em: $signingPfxPath"
    }
}
if ($signingPfxResolved) {
    $signtool = Get-Command signtool -ErrorAction SilentlyContinue
    if (-not $signtool) {
        Write-Host "Signtool nao encontrado no PATH. Instale o Windows SDK ou Build Tools."
    } else {
        $passwordValue = Get-SigningPassword -PasswordValue $signingPassword
        if ($passwordValue) {
            Invoke-SignFiles `
                -Paths @(
                    (Join-Path $distDir "cotacoes-moedas.exe")
                ) `
                -PfxPath $signingPfxResolved.Path `
                -Password $passwordValue `
                -TimestampUrl $signingTimestampUrl `
                -SignTool $signtool
        } else {
            Write-Host "Assinatura ignorada: senha do PFX nao informada."
        }
    }
} else {
    Write-Host "Assinatura desativada: defina SIGNING_PFX_PATH para assinar o exe e o instalador."
}

$isccPath = $null
if ($innoSetupPath) {
    $resolvedIscc = Resolve-Path $innoSetupPath -ErrorAction SilentlyContinue
    if ($resolvedIscc) {
        $isccPath = $resolvedIscc.Path
    } else {
        Write-Host "ISCC nao encontrado em: $innoSetupPath"
    }
}
if (-not $isccPath) {
    $isccCmd = Get-Command iscc -ErrorAction SilentlyContinue
    if ($isccCmd) {
        $isccPath = $isccCmd.Source
    }
}
if ($isccPath) {
    & $isccPath `
        (Join-Path $repoRoot "installer\\cotacoes-moedas.iss") `
        "/DMyAppName=$appName" `
        "/DMyAppPublisher=$appPublisher" `
        "/DMyAppSupportUrl=$appSupportUrl" `
        "/DMyAppDisplayVersion=$appVersion" `
        "/DMyAppVersion=$fileVersion" `
        "/DMyAppSourceDir=$distDir"

    $installerPath = Join-Path $repoRoot "dist\\cotacoes-moedas-setup.exe"
    if (Test-Path $installerPath) {
        Write-Host "Instalador gerado em: $installerPath"
    } else {
        Write-Host "Instalador nao encontrado em: $installerPath"
    }

    if ($signingPfxResolved -and $passwordValue -and $signtool) {
        Invoke-SignFiles `
            -Paths @((Join-Path $repoRoot "dist\\cotacoes-moedas-setup.exe")) `
            -PfxPath $signingPfxResolved.Path `
            -Password $passwordValue `
            -TimestampUrl $signingTimestampUrl `
            -SignTool $signtool
    }
} else {
    Write-Host "Inno Setup nao encontrado. Instale-o ou defina INNO_SETUP_PATH."
}
