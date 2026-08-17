# overTime

overTime is a PyQt6 desktop application for recording overtime, calculating
salary-based rates, and exporting clean Excel reports. Data remains on the
local computer in JSON files beside the application.

## Features

- Record overtime hours, dates, and task descriptions
- Calculate daily and hourly rates from a monthly salary
- Apply `x1`, `x1.5`, `x2`, or `x3` overtime multipliers
- Edit, sort, and remove entries directly in the application
- Export dated Excel reports with totals and formatted columns
- Persist data locally until it is reset by the user

## Screenshots

### Main Interface

![Overtime tracker main interface](https://github.com/user-attachments/assets/b47f7d74-9fbe-4d3a-a147-7ea47c3d9d14)

### Add an Entry

![Add overtime entry dialog](https://github.com/user-attachments/assets/cee062e5-2e62-4f87-8625-4a5a499ed8f4)

### Review and Edit Entries

![Editable overtime entries table](https://github.com/user-attachments/assets/cc3a35f5-de69-4b26-9ef3-b5a6f826add5)

### Excel Export

![Generated Excel overtime report](https://github.com/user-attachments/assets/cb6a1899-678c-4a4e-ac18-44d0883f071e)

## Run Locally

```powershell
git clone https://github.com/AhMadness/overTime.git
cd overTime
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

The app stores `salary_data.json` and `overtime_data.json` locally. Those files
are ignored by Git so personal salary and work records are not committed.

## Tests

```powershell
python -m unittest discover -s tests -v
```

## License

[MIT](LICENSE)
