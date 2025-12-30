# Freelance Dashboard

Excel KPI Dashboard for freelance activity tracking - Revenue, time tracking, clients with pivot tables, charts and VBA automation.

## Features

- **Structured Data Tables** - Clients, time entries, revenue tracking
- **Dynamic KPIs** - Total revenue, hourly rate, active clients, top client
- **Pivot Tables** - Revenue by client, by month, hours by project
- **Interactive Charts** - Pie chart, bar charts, trends
- **Slicers** - Filter by client and period
- **VBA Automation** - One-click refresh, rebuild dashboard
- **Professional Design** - Clean layout, conditional formatting

## File Structure

| Sheet | Content |
|-------|---------|
| Dashboard | Main view with KPIs, charts and slicers |
| Data_Clients | Client list (ID, name, sector, start date) |
| Data_Temps | Time entries (date, client, project, hours) |
| Data_Revenus | Revenue entries (date, client, amount, type) |
| Config | Settings (year, default rate, objectives) |
| TCD_Data | Pivot tables data |

## KPIs

- Total Revenue
- Current Month Revenue
- Total Hours
- Average Hourly Rate
- Number of Active Clients
- Top Client (by revenue)
- Hours This Week
- Unique Projects Count

## VBA Macros

| Macro | Description |
|-------|-------------|
| `RefreshDashboard` | Recalculates formulas and refreshes pivot tables |
| `QuickRefresh` | Silent refresh (no message) |
| `RebuildAll` | Rebuilds entire dashboard from scratch |
| `CreatePivotTables` | Creates/recreates pivot tables |
| `CreateCharts` | Creates/recreates charts |
| `CreateSlicers` | Creates/recreates slicers |
| `ApplyDesign` | Applies professional formatting |

## Screenshots

*Coming soon*

## Requirements

- Microsoft Excel 2016+ (or Microsoft 365)
- Macros enabled for VBA features

## Usage

1. Open `FreelanceDashboard.xlsm`
2. Enable macros when prompted
3. Add your data in Data_Clients, Data_Temps, Data_Revenus
4. Press `Alt+F8` and run `RefreshDashboard` to update

## Author

Alexis Trouve - alexistrouve.pro@gmail.com

## License

MIT License - See [LICENSE](LICENSE) file
