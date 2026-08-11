# 🎓 Canvas to Excel Organizer

Automatically fetch your Canvas LMS assignments and export them into a beautiful, color-coded Excel spreadsheet sorted by due date.


## ✨ Features

- **Automated Sync**: Pulls live assignment data directly from the Canvas API.
- **Smart Sorting**: Organizes tasks automatically by due date and course.
- **Visual Color-Coding**: Assigns a unique, custom color to each course for easy tracking.
- **Status Filtering**: (Optional) Filters out locked, completed, or ungraded assignments.

## 🚀 Getting Started

### Prerequisites

- Python 3.8 or higher
- A Canvas LMS account

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com
   cd canvas-to-excel
   ```

2. **Install required packages:**
   ```bash
   pip install -r requirements.txt
   ```

## ⚙️ Configuration

1. **Get your Canvas API Token:**
   - Log into Canvas.
   - Go to **Account** > **Settings**.
   - Scroll down to **Approved Integrations** and click **+ New Access Token**.
   - Copy the token immediately (you won't see it again).

2. **Set up environment variables:**
   Create a `.env` file in the root directory of the project and add your credentials:
   ```env
   CANVAS_API_URL="https://instructure.com"
   CANVAS_API_TOKEN="your_canvas_api_token_here"
   ```

## 💻 Usage

Run the main Python script to generate your spreadsheet:

```bash
python main.py
```

Once the script finishes, open the newly created `assignments.xlsx` file in the project folder to view your organized schedule.

## 📦 Built With

- [Canvasapi](https://github.com) - Python wrapper for the Canvas LMS API
- [Pandas](https://pydata.org) - Data manipulation and analysis
- [Openpyxl](https://readthedocs.io) - Excel spreadsheet formatting and styling

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
