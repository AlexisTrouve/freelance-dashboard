# Freelance Dashboard

Excel KPI Dashboard for freelance activity tracking - Revenue, time tracking, clients with dynamic formulas, pivot tables, charts and VBA automation.

## Features

- **Dynamic KPIs** - All metrics calculated with formulas (SUM, SUMIF, INDEX/MATCH)
- **Structured Data Tables** - tbl_Clients, tbl_Temps, tbl_Revenus
- **Pivot Tables** - Hours by client and project
- **Interactive Charts** - Pie chart (revenue), bar chart (hours)
- **Slicers** - Filter by client
- **Task Checklist** - Interactive checkboxes with linked cells
- **VBA Automation** - One-click refresh

## Dashboard Preview

| KPI | Formula |
|-----|---------|
| Total Revenue | `=SUM(tbl_Revenus[Montant])` |
| Total Hours | `=SUM(tbl_Temps[Heures])` |
| Hourly Rate | `=ROUND(CA/Hours,2)` |
| Client Count | `=COUNTA(tbl_Clients[ClientID])` |
| Top Client | `=INDEX(...MATCH(MAX(...)))` |

## File Structure

| Sheet | Content |
|-------|---------|
| Dashboard | KPIs, charts, pivot table, slicer, task list |
| Data_Clients | Client list (ID, name, sector, start date) |
| Data_Temps | Time entries (date, client, project, hours) |
| Data_Revenus | Revenue entries (date, client, amount, type) |

## Requirements

- Microsoft Excel 2016+ (or Microsoft 365)
- Macros enabled for VBA features

## Usage

1. Open `templates/FreelanceDashboard.xlsm`
2. Enable macros when prompted
3. Add your data in the Data_* sheets
4. Dashboard updates automatically (or run `RefreshDashboard` macro)

## Author

Alexis Trouve - alexistrouve.pro@gmail.com

## License

MIT License - See [LICENSE](LICENSE) file
