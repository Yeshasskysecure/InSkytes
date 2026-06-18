# SharePoint Migration Site - Setup Script
# This script helps set up the SPFx development environment

Write-Host "SharePoint Migration Site - Setup Script" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green
Write-Host ""

# Check Node.js
Write-Host "Checking Node.js installation..." -ForegroundColor Yellow
try {
    $nodeVersion = node --version
    Write-Host "Node.js found: $nodeVersion" -ForegroundColor Green
    
    # Check if version is 18 or higher
    $majorVersion = [int]($nodeVersion -replace 'v(\d+)\..*', '$1')
    if ($majorVersion -lt 18) {
        Write-Host "Warning: Node.js version 18 or higher is recommended for SPFx" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Error: Node.js is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install Node.js from https://nodejs.org/" -ForegroundColor Yellow
    exit 1
}

# Check npm
Write-Host "Checking npm installation..." -ForegroundColor Yellow
try {
    $npmVersion = npm --version
    Write-Host "npm found: $npmVersion" -ForegroundColor Green
} catch {
    Write-Host "Error: npm is not installed or not in PATH" -ForegroundColor Red
    Write-Host "npm should come with Node.js installation" -ForegroundColor Yellow
    exit 1
}

# Check Gulp CLI
Write-Host "Checking Gulp CLI..." -ForegroundColor Yellow
try {
    $gulpVersion = gulp --version
    Write-Host "Gulp CLI found" -ForegroundColor Green
} catch {
    Write-Host "Gulp CLI not found globally. Installing..." -ForegroundColor Yellow
    npm install -g gulp-cli
}

# Install dependencies
Write-Host ""
Write-Host "Installing project dependencies..." -ForegroundColor Yellow
Write-Host "This may take several minutes..." -ForegroundColor Yellow
npm install

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "Setup completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. Build the solution: npm run build" -ForegroundColor White
    Write-Host "2. Serve locally: gulp serve" -ForegroundColor White
    Write-Host "3. Open the workbench at: https://indegenedevelopment.sharepoint.com/sites/KnowledgeHub/_layouts/workbench.aspx" -ForegroundColor White
} else {
    Write-Host ""
    Write-Host "Setup failed. Please check the error messages above." -ForegroundColor Red
    exit 1
}

