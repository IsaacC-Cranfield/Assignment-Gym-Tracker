from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.trendline import Trendline
from collections import defaultdict, Counter
from datetime import datetime
import numpy as np
import os

# Define workouts
workouts = {
    "Back & Biceps": ["Pull-ups", "Chest Supported Row", "Lat Pulldown", "Preacher Bicep Curl", "Hammer Curls"],
    "Chest & Triceps": ["Smith Machine Incline Press", "Incline Dumbbell Press", "Flat Bench Machine", "Tricep Pushdown", "Dips"],
    "Leg Day": ["Squats", "Leg Press", "Leg Curls", "Quad Extensions", "Calf Raises"],
    "Shoulders": ["Shoulder Press", "Lateral Raises", "Rear Delt Flyes"]
}

filename = "gym_progress1.xlsx"

# Load or create workbook
if os.path.exists(filename):
    wb = load_workbook(filename)
else:
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet if new

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

# Step 3: Input sets
num_sets = int(input(f"\nHow many sets for {selected_workout}? "))
now = datetime.now()
date_str = now.strftime("%Y-%m-%d %H:%M")
week_number = now.strftime("%Y-W%U")

# Use sheet named after the workout
sheet_name = selected_workout
if sheet_name not in wb.sheetnames:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(["Date", "Week", "Training Day", "Workout", "Set #", "Weight (kg)", "Reps", "Volume", "1RM Estimate"])
else:
    ws = wb[sheet_name]

for set_num in range(1, num_sets + 1):
    weight = float(input(f"Enter weight for set {set_num} (kg): "))
    reps = int(input(f"Enter reps for set {set_num}: "))
    volume = weight * reps
    one_rm = round(weight * (1 + reps / 30), 2)
    ws.append([date_str, week_number, selected_day, selected_workout, set_num, weight, reps, volume, one_rm])

#creates 1 rep max chart
def create_weekly_1rm_chart():
    ws_data = wb[selected_workout]  # same sheet as workout data

    # Find the starting row for writing the summary
    start_row = ws_data.max_row + 3  # a few rows after the log

    # Create a dictionary of weekly 1RM estimates
    weekly_1rms = defaultdict(list)
    for row in ws_data.iter_rows(min_row=2, values_only=True):
        _, week, _, _, _, _, _, _, one_rm = row
        if one_rm:
            weekly_1rms[week].append(one_rm)

    # Write weekly averages below existing data
    ws_data.cell(row=start_row, column=1, value="Week")
    ws_data.cell(row=start_row, column=2, value="Avg 1RM Estimate")

    for i, (week, values) in enumerate(sorted(weekly_1rms.items()), start=start_row + 1):
        avg_rm = round(sum(values) / len(values), 2)
        ws_data.cell(row=i, column=1, value=week)
        ws_data.cell(row=i, column=2, value=avg_rm)

    # Create the chart itself
    chart = LineChart()
    chart.title = f"{selected_workout} – Weekly Avg 1RM Estimate"
    chart.y_axis.title = "1RM (kg)"
    chart.x_axis.title = "Week"

    data = Reference(ws_data, min_col=2, min_row=start_row, max_row=i)
    categories = Reference(ws_data, min_col=1, min_row=start_row + 1, max_row=i)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    # Place the chart somewhere visually clear
    chart_position = f"E{start_row}"
    ws_data.add_chart(chart, chart_position)
    
# Forecasting
def create_progress_forecast_chart(target_reps=12):
    sheet_title = f"{selected_workout}_Progress"
    if sheet_title not in wb.sheetnames:
        ws_chart = wb.create_sheet(title=sheet_title)
        ws_chart.append(["Week", "Max Weight × Reps"])
    else:
        ws_chart = wb[sheet_title]

    weekly_progress = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        week = row[1]
        weight = row[5]
        reps = row[6]

        # ✅ Skip rows with missing data
        if week is None or weight is None or reps is None:
            continue

        score = weight * reps
        if week not in weekly_progress or score > weekly_progress[week]:
            weekly_progress[week] = score

    existing_weeks = {row[0] for row in ws_chart.iter_rows(min_row=2, values_only=True)}
    for week, score in sorted(weekly_progress.items()):
        if week not in existing_weeks:
            ws_chart.append([week, round(score, 2)])

    # Chart creation code continues...


    # Create line chart
    chart = LineChart()
    chart.title = f"{selected_workout} – Weekly Progress"
    chart.x_axis.title = "Week"
    chart.y_axis.title = "Best Weight × Reps"
    data = Reference(ws_chart, min_col=2, min_row=1, max_row=ws_chart.max_row)
    cats = Reference(ws_chart, min_col=1, min_row=2, max_row=ws_chart.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].trendline = Trendline(trendlineType='linear')
    ws_chart.add_chart(chart, "E2")

    # Estimate future milestone 1:1? ;)
    weeks = []
    values = []
    for i, row in enumerate(ws_chart.iter_rows(min_row=2, values_only=True), start=1):
        weeks.append(i)
        values.append(row[1])
    if len(weeks) >= 2:
        x = np.array(weeks)
        y = np.array(values)
        a, b = np.polyfit(x, y, 1)
        latest_weight = max(row[5] for row in ws.iter_rows(min_row=2, values_only=True))
        target_volume = latest_weight * target_reps
        if a > 0:
            est_week = int((target_volume - b) / a)
            print(f"\n📈 Estimated to hit {target_volume:.0f} (e.g., {latest_weight:.0f}kg × {target_reps} reps) around week {est_week}.")
        else:
            print("\n⚠️ Progress is flat or decreasing.")
    else:
        print("\nℹ️ Not enough data to forecast.")
        

# Run processes
create_weekly_1rm_chart()
create_progress_forecast_chart()
wb.save(filename)
print(f"\n✅ Workout saved to '{filename}' with per-workout tracking.")

# Are you not entertained? Gladiator reference.
