# Assignment-Gym-Tracker
from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from datetime import datetime
import os

# Define workouts by training day
workouts = {
    "Back & Biceps": ["Pull-ups", "Barbell Row", "Lat Pulldown", "Bicep Curl"],
    "Chest & Triceps": ["Bench Press", "Incline Dumbbell Press", "Tricep Pushdown", "Dips"],
    "Leg Day": ["Squats", "Leg Press", "Lunges", "Calf Raises"],
    "Shoulders": ["Shoulder Press", "Lateral Raises", "Front Raises"]
}

# Excel file name
filename = "gym_progress.xlsx"

# Load or create workbook
if os.path.exists(filename):
    wb = load_workbook(filename)
    ws = wb.active
else:
    wb = Workbook()
    ws = wb.active
    ws.title = "Progress"
    ws.append(["Date", "Training Day", "Workout", "Weight (kg)"])  # Header

# Step 1: Select training day
print("Select Training Day:")
training_days = list(workouts.keys())
for idx, day in enumerate(training_days):
    print(f"{idx + 1}. {day}")
day_choice = int(input("Enter your choice: ")) - 1
selected_day = training_days[day_choice]

# Step 2: Select workout
print(f"\nSelect Workout for {selected_day}:")
selected_workouts = workouts[selected_day]
for idx, workout in enumerate(selected_workouts):
    print(f"{idx + 1}. {workout}")
workout_choice = int(input("Enter your choice: ")) - 1
selected_workout = selected_workouts[workout_choice]

# Step 3: Enter weight used
weight = float(input(f"Enter weight used for {selected_workout} (kg): "))

# Step 4: Save to Excel
date_now = datetime.now().strftime("%Y-%m-%d %H:%M")
ws.append([date_now, selected_day, selected_workout, weight])

# Step 5: Create a chart for the selected workout
def create_chart():
    # Find all rows with the selected workout
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    workout_data = [(r[0], r[3]) for r in rows[1:] if r[2] == selected_workout]

    if len(workout_data) < 2:
        print("Not enough data points to create a chart yet.")
        return

    # Write workout-specific data to a new sheet
    if selected_workout not in wb.sheetnames:
        ws_chart = wb.create_sheet(title=selected_workout)
        ws_chart.append(["Date", "Weight (kg)"])
    else:
        ws_chart = wb[selected_workout]

    # Add only new entries
    existing_dates = {row[0] for row in ws_chart.iter_rows(min_row=2, values_only=True)}
    for date, weight_value in workout_data:
        if date not in existing_dates:
            ws_chart.append([date, weight_value])

    # Create chart
    chart = LineChart()
    chart.title = f"{selected_workout} Progress"
    chart.y_axis.title = "Weight (kg)"
    chart.x_axis.title = "Date"

    data = Reference(ws_chart, min_col=2, min_row=1, max_row=ws_chart.max_row)
    categories = Reference(ws_chart, min_col=1, min_row=2, max_row=ws_chart.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    ws_chart.add_chart(chart, "E2")

# Generate chart
create_chart()

# Save workbook
wb.save(filename)
print(f"\nWorkout logged and chart updated in '{filename}'.")
