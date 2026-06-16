# Financial Tracker

A web-based financial tracker that runs entirely on a Windows PC — **zero external dependencies**. Uses PowerShell as the HTTP server and Excel as the database.

## Requirements (all pre-installed on most Windows PCs)

- Windows (any edition)
- PowerShell 5.1+
- Microsoft Excel

## Quick Start

1. **Copy** the `financial-tracker` folder to your Windows PC
2. **Open PowerShell** and navigate to the folder:
   ```powershell
   cd C:\path\to\financial-tracker
   ```
3. **Set execution policy** (first time only):
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   ```
4. **Run the server:**
   ```powershell
   .\server.ps1
   ```
5. **Open your browser** and go to: `http://localhost:8080`

## Features

- **Dashboard**: Balance, total income, total expense, transaction count
- **Transaction Management**: Add, edit, delete transactions
- **Search & Filter**: Filter by type (Income/Expense), category, or keyword search
- **Category Breakdown**: Visual category analysis with income/expense bars
- **Excel Backend**: All data stored in `data.xlsx` — openable directly in Excel
- **Dark Theme**: Modern dark UI

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/transactions` | List all transactions (supports `?search=`, `?type=`, `?category=`) |
| POST | `/api/transactions` | Add new transaction |
| PUT | `/api/transactions/{id}` | Update transaction |
| DELETE | `/api/transactions/{id}` | Delete transaction |
| GET | `/api/summary` | Get summary stats and category breakdown |

## Data File

Your transactions are stored in `data.xlsx`. You can open it directly in Excel to view, edit, or create reports. The file is created automatically on first run.

## File Structure

```
financial-tracker/
├── server.ps1      # PowerShell HTTP server + Excel COM backend
├── index.html      # Web frontend (HTML/CSS/JS — single file, no dependencies)
├── data.xlsx       # Auto-created Excel database (appears on first run)
└── README.md       # This file
```
