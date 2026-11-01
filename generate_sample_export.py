"""
Generate a synthetic Rally-style sprint export for practice.
All names, IDs and numbers are invented. No real project data.

Run on your Mac:
    python3 generate_sample_export.py
Output:
    sample_rally_export.xlsx  (sheets: Stories, Iterations, Data Dictionary)
"""
import datetime as dt
import random

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

random.seed(26)  # same seed = same sample every time you run it

TODAY = dt.date(2026, 10, 31)
STATES = ["Defined", "In-Progress", "Completed", "Accepted"]
OWNERS = ["Owner A", "Owner B", "Owner C", "Owner D", "Owner E", "Owner F"]
POINTS = [1, 2, 3, 5, 8]
BLOCK_REASONS = [
    "Waiting on vendor API access",
    "Dependency on access review sign-off",
    "Test environment not ready",
    "Decision pending from steering group",
]
VERBS = ["Add", "Update", "Refine", "Automate", "Review", "Fix", "Document", "Test"]
OBJECTS = [
    "approval step", "status report", "data validation", "notification rule",
    "export format", "filter view", "audit trail", "dashboard widget",
]
AREAS = [
    "for patch scheduling", "for asset inventory", "for access reviews",
    "for the reporting portal", "for vendor onboarding",
]
COLUMNS = [
    "Formatted ID", "Name", "Schedule State", "Iteration", "Plan Estimate",
    "Owner", "Blocked", "Blocked Reason", "Creation Date", "Accepted Date", "Project",
]
PROJECT = "Sample Project"

# Seven two-week sprints starting 3 Aug 2026. Sprint 07 is the current sprint.
iterations = []
for i in range(7):
    start = dt.date(2026, 8, 3) + dt.timedelta(days=14 * i)
    end = start + dt.timedelta(days=11)
    iterations.append((f"Sprint {i + 1:02d}", start, end))

stories = []
story_no = 1001
for name, start, end in iterations:
    is_current = start <= TODAY <= end
    is_closed = end < TODAY
    for _ in range(random.randint(9, 12)):
        if is_current:
            state = random.choices(STATES, weights=[25, 40, 10, 25])[0]
        elif is_closed:
            state = random.choices(STATES, weights=[0, 15, 10, 75])[0]
        else:
            state = "Defined"

        blocked, reason = "No", ""
        if state in ("Defined", "In-Progress") and random.random() < 0.15:
            blocked, reason = "Yes", random.choice(BLOCK_REASONS)

        accepted = ""
        if state == "Accepted":
            accepted_date = end - dt.timedelta(days=random.randint(0, 4))
            accepted = min(accepted_date, TODAY)

        creation = start - dt.timedelta(days=random.randint(3, 20))
        story_name = (
            f"{random.choice(VERBS)} {random.choice(OBJECTS)} {random.choice(AREAS)}"
        )
        stories.append([
            f"US{story_no}",
            story_name,
            state,
            name,
            random.choice(POINTS),
            random.choice(OWNERS),
            blocked,
            reason,
            creation,
            accepted,
            PROJECT,
        ])
        story_no += 1

wb = Workbook()
header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor="1F4E78")


def style_header(ws, ncols):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def set_widths(ws, widths):
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width


# Sheet 1: Stories (the main export)
ws = wb.active
ws.title = "Stories"
ws.append(COLUMNS)
for row in stories:
    ws.append(row)
style_header(ws, len(COLUMNS))
for row in ws.iter_rows(min_row=2):
    row[8].number_format = "yyyy-mm-dd"   # Creation Date
    row[9].number_format = "yyyy-mm-dd"   # Accepted Date
ws.auto_filter.ref = ws.dimensions
set_widths(ws, [14, 60, 16, 12, 14, 12, 10, 36, 14, 14, 16])

# Sheet 2: Iterations (one row per sprint)
wi = wb.create_sheet("Iterations")
wi.append(["Iteration", "Start Date", "End Date", "State", "Project"])
for name, start, end in iterations:
    if end < TODAY:
        state = "Accepted"
    elif start <= TODAY <= end:
        state = "Committed"
    else:
        state = "Planning"
    wi.append([name, start, end, state, PROJECT])
style_header(wi, 5)
for row in wi.iter_rows(min_row=2):
    row[1].number_format = "yyyy-mm-dd"
    row[2].number_format = "yyyy-mm-dd"
set_widths(wi, [14, 14, 14, 12, 16])

# Sheet 3: Data Dictionary (what each column means)
wd = wb.create_sheet("Data Dictionary")
wd.append(["Column", "Meaning", "Example", "Used for"])
dictionary = [
    ("Formatted ID", "Unique story ID", "US1001", "Counting, tracing back to Rally"),
    ("Name", "Story title", "Add audit trail for access reviews", "Tables"),
    ("Schedule State", "Workflow state: Defined, In-Progress, Completed, Accepted", "Accepted", "Status chart"),
    ("Iteration", "Sprint the story is planned in", "Sprint 06", "Grouping by sprint"),
    ("Plan Estimate", "Story points committed", "5", "Committed vs completed"),
    ("Owner", "Person responsible", "Owner A", "Workload view"),
    ("Blocked", "Yes or No", "Yes", "Blocker table"),
    ("Blocked Reason", "Why the story is blocked", "Waiting on vendor API access", "Blocker table"),
    ("Creation Date", "When the story was created", "2026-09-28", "Ageing"),
    ("Accepted Date", "When the story was accepted (blank if not accepted)", "2026-10-22", "Throughput"),
    ("Project", "Project name", "Sample Project", "Filter"),
]
for row in dictionary:
    wd.append(list(row))
style_header(wd, 4)
set_widths(wd, [18, 58, 36, 30])

wb.save("sample_rally_export.xlsx")
print(f"Created sample_rally_export.xlsx with {len(stories)} stories across {len(iterations)} sprints")
