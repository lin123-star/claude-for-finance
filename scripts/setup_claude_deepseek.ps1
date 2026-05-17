param(
    [ValidateSet('deepseek-v4-pro','deepseek-v4-flash')]
    [string]$Model = 'deepseek-v4-pro',
    [string]$ApiKey,
    [switch]$SkipConnectivityTest
)

$ErrorActionPreference = 'Stop'

function Remove-EnvVar {
    param([string]$Name)

    Remove-Item "Env:$Name" -ErrorAction SilentlyContinue
    [Environment]::SetEnvironmentVariable($Name, $null, 'User')
}

Write-Host 'Step 1/4: Cleaning old provider variables...' -ForegroundColor Cyan
$varsToClean = @(
    'KIMI_API_KEY','MOONSHOT_API_KEY','MOONSHOT_BASE_URL',
    'OPENAI_API_KEY','OPENAI_BASE_URL',
    'ANTHROPIC_API_KEY','ANTHROPIC_BASE_URL','ANTHROPIC_MODEL',
    'DEEPSEEK_API_KEY','DEEPSEEK_BASE_URL'
)

foreach ($v in $varsToClean) {
    Remove-EnvVar -Name $v
}

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    Write-Host 'Step 2/4: Reading DeepSeek API key (input hidden)...' -ForegroundColor Cyan
    $secure = Read-Host 'Enter DeepSeek API key' -AsSecureString
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        $deepseekKey = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}
else {
    Write-Host 'Step 2/4: Using API key from parameter...' -ForegroundColor Cyan
    $deepseekKey = $ApiKey
}

if ([string]::IsNullOrWhiteSpace($deepseekKey)) {
    throw 'DeepSeek API key cannot be empty.'
}

Write-Host 'Step 3/4: Writing Claude-compatible env vars for DeepSeek...' -ForegroundColor Cyan
$anthropicBase = 'https://api.deepseek.com/anthropic'

$env:ANTHROPIC_BASE_URL = $anthropicBase
$env:ANTHROPIC_API_KEY = $deepseekKey
$env:ANTHROPIC_MODEL = $Model

[Environment]::SetEnvironmentVariable('ANTHROPIC_BASE_URL', $anthropicBase, 'User')
[Environment]::SetEnvironmentVariable('ANTHROPIC_API_KEY', $deepseekKey, 'User')
[Environment]::SetEnvironmentVariable('ANTHROPIC_MODEL', $Model, 'User')

Write-Host "Saved: ANTHROPIC_BASE_URL=$anthropicBase" -ForegroundColor Green
Write-Host "Saved: ANTHROPIC_MODEL=$Model" -ForegroundColor Green
Write-Host 'Saved: ANTHROPIC_API_KEY=***hidden***' -ForegroundColor Green

if (-not $SkipConnectivityTest) {
    Write-Host 'Step 4/4: Testing Anthropic-compatible endpoint...' -ForegroundColor Cyan

    $body = @{
        model = $Model
        max_tokens = 64
        messages = @(
            @{ role = 'user'; content = 'Reply exactly: CONNECT_OK' }
        )
    } | ConvertTo-Json -Depth 8

    $requestParams = @{
        Method  = 'Post'
        Uri     = 'https://api.deepseek.com/anthropic/v1/messages'
        Headers = @{
            'x-api-key'          = $deepseekKey
            'anthropic-version'  = '2023-06-01'
            'content-type'       = 'application/json'
        }
        Body    = $body
    }

    $resp = Invoke-RestMethod @requestParams

    Write-Host 'Connectivity test succeeded.' -ForegroundColor Green
    if ($resp.content) {
        $text = ($resp.content | Where-Object { $_.type -eq 'text' } | Select-Object -First 1).text
        if ($text) {
            Write-Host ("Model reply: " + $text)
        }
    }
}
else {
    Write-Host 'Step 4/4: Connectivity test skipped.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host 'Done. Next actions:' -ForegroundColor Green
Write-Host '1) Fully restart VS Code.'
Write-Host '2) Open a new terminal and run: claude'
Write-Host '3) If still wrong provider, run: Get-ChildItem Env: | ? Name -match "KIMI|MOONSHOT|OPENAI|ANTHROPIC|DEEPSEEK"'
