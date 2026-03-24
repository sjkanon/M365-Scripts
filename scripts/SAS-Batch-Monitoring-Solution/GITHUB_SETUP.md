# GitHub Setup Guide

## Quick Start

1. **Create a new GitHub repository:**
   ```bash
   # On GitHub.com, create new repo: sas-batch-monitoring
   ```

2. **Extract and push the files:**
   ```bash
   # Extract the zip
   unzip SAS-Batch-Monitoring-Solution.zip -d sas-batch-monitoring
   cd sas-batch-monitoring
   
   # Initialize git repo
   git init
   git add .
   git commit -m "Initial commit: SAS Batch Monitoring Solution"
   
   # Link to GitHub
   git remote add origin https://github.com/yourusername/sas-batch-monitoring.git
   git branch -M main
   git push -u origin main
   ```

## Recommended Repository Structure

```
sas-batch-monitoring/
├── README.md                              # Main documentation
├── EVENTLOG_GUIDE.md                      # Event Viewer integration guide
├── LICENSE                                # MIT License
├── .gitignore                             # Git ignore patterns
├── scripts/
│   ├── Monitor-SASBatchErrors.ps1         # Basic monitoring
│   ├── Monitor-SASBatchErrors-Enhanced.ps1 # With Event Viewer
│   ├── Test-SASWorkDirectory.ps1          # Disk health check
│   └── Setup-SASMonitoring.ps1            # Installation script
├── config/
│   └── zabbix_sas_monitor.conf            # Zabbix configuration
├── examples/
│   └── sample_output.txt                  # Example outputs
└── docs/
    └── troubleshooting.md                 # Additional docs
```

## Repository Settings

### Description
```
Automated monitoring solution for SAS batch jobs with Event Viewer integration and Zabbix support
```

### Topics (tags)
- `sas`
- `monitoring`
- `powershell`
- `zabbix`
- `event-viewer`
- `batch-jobs`
- `windows-server`
- `error-detection`

### About Section
```
🔍 Comprehensive SAS batch job monitoring
✅ Event Viewer integration
📊 Zabbix support
🚨 Email alerts
📝 Detailed error reporting
```

## README Badges

Add these to your README.md:

```markdown
![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows%20Server-lightgrey.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)
```

## .github/workflows (Optional CI)

Create `.github/workflows/powershell-lint.yml`:

```yaml
name: PowerShell Linting

on: [push, pull_request]

jobs:
  lint:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      - name: Run PSScriptAnalyzer
        shell: pwsh
        run: |
          Install-Module -Name PSScriptAnalyzer -Force -Scope CurrentUser
          Invoke-ScriptAnalyzer -Path . -Recurse -Settings PSGallery
```

## Release Notes Template

When creating releases, use this format:

```markdown
## Version 1.0.0 - Initial Release

### Features
- ✨ SAS log file error detection
- ✨ Windows Event Viewer integration
- ✨ Zabbix monitoring support
- ✨ Email alert configuration
- ✨ Automated installation script

### Error Detection
- Can't spawn sas.bat failures
- WORK library authorization errors
- Disk/storage system errors
- SAS service crashes
- Permission/access denied events

### Installation
Download and run:
```powershell
.\Setup-SASMonitoring.ps1 -InstallZabbix
```

### Documentation
- [README.md](README.md) - Main documentation
- [EVENTLOG_GUIDE.md](EVENTLOG_GUIDE.md) - Event Viewer integration

### Tested On
- Windows Server 2019/2022
- PowerShell 5.1+
- Zabbix Agent 2
```

## Protect Main Branch

In GitHub repository settings:
1. Go to Settings → Branches
2. Add branch protection rule for `main`
3. Enable:
   - ✅ Require pull request reviews
   - ✅ Require status checks to pass
   - ✅ Require branches to be up to date

## Issues Template

Create `.github/ISSUE_TEMPLATE/bug_report.md`:

```markdown
---
name: Bug Report
about: Report an issue with the monitoring scripts
---

**Script Name:**
Which script has the issue? (Monitor-SASBatchErrors.ps1, Test-SASWorkDirectory.ps1, etc.)

**Environment:**
- Windows Version:
- PowerShell Version:
- Zabbix Agent Version (if applicable):

**Expected Behavior:**
What should happen?

**Actual Behavior:**
What actually happens?

**Error Output:**
```
Paste error messages here
```

**Steps to Reproduce:**
1. 
2. 
3. 

**Additional Context:**
Add any other context about the problem here.
```

## Star & Watch

Remember to star your own repo and watch it for issues/PRs! 🌟
