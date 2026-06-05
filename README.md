
```markdown
# QA Audit CSV to Excel Dashboard Converter

This Python automation script takes raw QA audit logs (in CSV format) and converts them into a formatted, presentation-ready Excel file (.xlsx). It features conditional formatting to highlight different AI decisions and automatically generates a summary dashboard with a pie chart.

## Features
* **Automated Data Transfer:** Converts CSV rows into Excel seamlessly.
* **Conditional Formatting:** Color-codes cells based on specific text values (e.g., Green for `NEW_DAMAGE`, Yellow for `DISCARD`).
* **Auto-Sizing:** Adjusts column widths and sets text wrapping for long text descriptions.
* **Dashboard Generation:** Automatically calculates value counts and creates a Pie Chart in a separate sheet for quick presentations.

## Prerequisites

Make sure you have Python 3.x installed on your system. 

## Installation

1. Clone this repository or download the source code:
   ```bash
   git clone https://github.com/YourUsername/QA-Audit-Excel-Tool.git
   cd QA-Audit-Excel-Tool

```

2. Install the required Python dependencies using `pip`:
```bash
pip install -r requirements.txt

```



## Usage

1. Place your data file named `qa_audit_tracker.csv` in the same directory as the script. The script expects the CSV to have a column named `AI Decision` in the 4th position.
2. Run the script:
```bash
python main.py

```


3. A new file named `QA_Audit_Presentation.xlsx` will be generated in the same folder, containing your formatted data and interactive dashboard.

## Customization

You can easily customize the tool by editing the `main.py` file:

* **Change Input/Output names:** Modify the `INPUT_FILE` and `OUTPUT_FILE` variables at the bottom of the script.
* **Change Colors:** Update the hex codes inside the `colors` dictionary to match your specific needs.

## License

This project is open-source and available under the MIT License.
