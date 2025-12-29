# Dashboard Automation

Automatically updates the Associate Progression Dashboard every Saturday at noon.

## Overview

This repository contains the automation script and GitHub Actions workflow to keep the dashboard current with the latest associate data.

## How It Works

- **Schedule**: Every Saturday at 12:00 PM UTC (4:00 PM PST)
- **Process**:
  1. Extract latest week data from Excel files
  2. Update master data file
  3. Regenerate chunk files
  4. Upload to Netlify
  5. Dashboard updates live

## Files

- `DASHBOARD_AUTO_DEPLOY.py` - Main automation script
- `.github/workflows/deploy.yml` - GitHub Actions configuration
- `README.md` - This file

## Dashboard

https://associateprogressiondashboard.netlify.app

## Secrets

GitHub Secrets (configured in Settings → Secrets):
- `NETLIFY_TOKEN` - Netlify API token
- `NETLIFY_SITE_ID` - Netlify site ID

## Monitoring

Check automation runs in the **Actions** tab of this repository.

