from ortools.sat.python import cp_model

import csv
from pathlib import Path
from typing import Union

import class_config as cfg

model = cp_model.CpModel()

def _normalize_sex(value: str) -> str:
    v = (value or "").strip().lower()
    if v in {"k", "f", "female", "kvinde"}:
        return "F"
    if v in {"m", "male", "mand"}:
        return "M"
    # Fall back to original (uppercased) if unknown
    return (value or "").strip().upper()


def load_students_from_csv(csv_path: Union[str, Path]):
    """
    Expected CSV columns (0-indexed):
      0: sex
      1: id
      2: language
      3: subject
      4: origin
    """
    csv_path = Path(csv_path)
    students_local = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row or all(not (c or "").strip() for c in row):
                continue
            if len(row) < 5:
                continue

            sex_raw, id_raw, language, subject, origin = row[:5]

            # Skip header rows (e.g. "Køn,Navn,...") by checking id is int-like
            try:
                source_id = int(str(id_raw).strip())
            except ValueError:
                continue

            internal_id = len(students_local)
            students_local.append(
                {
                    "id": internal_id,  # solver-friendly sequential id
                    "source_id": source_id,  # id from CSV
                    "name": str(source_id),
                    "sex": _normalize_sex(sex_raw),
                    "language": (language or "").strip(),
                    "subject": (subject or "").strip(),
                    "origin": (origin or "").strip(),
                }
            )
    return students_local


students = load_students_from_csv(cfg.CSV_PATH)
classes = cfg.CLASSES

assignments = {}
for s in students:
    for c in classes:
        assignments[(s["id"], c)] = model.NewBoolVar(f's{s["id"]}_in_{c}')

# Each student must be assigned to exactly one class
for s in students:
    model.Add(sum(assignments[(s["id"], c)] for c in classes) == 1)

min_female = cfg.MIN_FEMALE_PER_CLASS
max_female = len(students) - 2
min_male = cfg.MIN_MALE_PER_CLASS
max_male = len(students) - 2
for c in classes:
    # Sum only variables where the student's sex is 'F'
    female_count = sum(
        assignments[(s["id"], c)] 
        for s in students if s["sex"] == "F"
    )
    male_count = sum(
        assignments[(s["id"], c)] 
        for s in students if s["sex"] == "M"
    )
    model.Add(female_count >= min_female)
    model.Add(female_count <= max_female)
    model.Add(male_count >= min_male)
    model.Add(male_count <= max_male)
# Hard constraint: class capacity
MAX_CLASS_SIZE = cfg.MAX_CLASS_SIZE
MIN_CLASS_SIZE = cfg.MIN_CLASS_SIZE
class_sizes = {}
for c in classes:
    class_size = model.NewIntVar(0, len(students), f"class_size_{c}")
    model.Add(class_size == sum(assignments[(s["id"], c)] for s in students))
    class_sizes[c] = class_size
    model.Add(class_size <= MAX_CLASS_SIZE)
    model.Add(class_size >= MIN_CLASS_SIZE)

penalties = []

# Soft rule: balance class sizes around the average.
# Penalize absolute deviation from the closest integer target (floor/ceil of N/K).
CLASS_BALANCE_WEIGHT = cfg.CLASS_BALANCE_WEIGHT
target_floor = len(students) // len(classes)
target_ceil = (len(students) + len(classes) - 1) // len(classes)
for c in classes:
    dev_floor = model.NewIntVar(0, len(students), f"dev_floor_{c}")
    dev_ceil = model.NewIntVar(0, len(students), f"dev_ceil_{c}")
    model.AddAbsEquality(dev_floor, class_sizes[c] - target_floor)
    model.AddAbsEquality(dev_ceil, class_sizes[c] - target_ceil)
    dev = model.NewIntVar(0, len(students), f"size_dev_{c}")
    model.AddMinEquality(dev, [dev_floor, dev_ceil])
    penalties.append(dev * CLASS_BALANCE_WEIGHT)

ORIGIN_WEIGHT = cfg.ORIGIN_WEIGHT
origins = sorted({s["origin"] for s in students if (s.get("origin") or "").strip()})
for c in classes:
    for origin in origins:
        # Count students from this origin in this class
        origin_count = sum(
            assignments[(s["id"], c)] 
            for s in students if s["origin"] == origin
        )
        
        # Create a 'excess' variable: how many students over the per-origin limit?
        excess = model.NewIntVar(0, len(students), f'excess_{origin}_{c}')
        
        # excess >= (origin_count - ORIGIN_MAX_PER_CLASS)
        model.Add(excess >= origin_count - cfg.ORIGIN_MAX_PER_CLASS)
        
        # We want to minimize this excess
        penalties.append(excess * ORIGIN_WEIGHT)  # weight/cost of this penalty

# Soft rule: prefer students with the same subject to be in the same class.
# Penalty = weight * (number of classes that have this subject - 1), so 0 when all are together.
SUBJECT_WEIGHT = cfg.SUBJECT_WEIGHT
subjects = list({s["subject"] for s in students})
for subject in subjects:
    # For each class, indicate whether it has any student from this subject
    has_in_class = []
    for c in classes:
        subject_count = sum(
            assignments[(s["id"], c)] for s in students if s["subject"] == subject
        )
        has = model.NewBoolVar(f"has_{subject}_{c}")
        model.Add(subject_count >= 1).OnlyEnforceIf(has)
        model.Add(subject_count == 0).OnlyEnforceIf(has.Not())
        has_in_class.append(has)
    # Spread = (number of classes with this subject) - 1; 0 when all in one class
    num_classes_with = model.NewIntVar(0, len(classes), f"nclass_{subject}")
    model.Add(num_classes_with == sum(has_in_class))
    spread = model.NewIntVar(0, len(classes) - 1, f"spread_{subject}")
    model.Add(spread == num_classes_with - 1)
    penalties.append(spread * SUBJECT_WEIGHT)

# Soft rule: prefer students with the same language to be in the same class.
# Penalty = weight * (number of classes that have this language - 1), so 0 when all are together.
LANGUAGE_WEIGHT = cfg.LANGUAGE_WEIGHT
languages = list({s["language"] for s in students if (s.get("language") or "").strip()})
for language in languages:
    has_in_class = []
    for c in classes:
        language_count = sum(
            assignments[(s["id"], c)] for s in students if s["language"] == language
        )
        has = model.NewBoolVar(f"has_lang_{language}_{c}")
        model.Add(language_count >= 1).OnlyEnforceIf(has)
        model.Add(language_count == 0).OnlyEnforceIf(has.Not())
        has_in_class.append(has)
    num_classes_with = model.NewIntVar(0, len(classes), f"nclass_lang_{language}")
    model.Add(num_classes_with == sum(has_in_class))
    spread = model.NewIntVar(0, len(classes) - 1, f"spread_lang_{language}")
    model.Add(spread == num_classes_with - 1)
    penalties.append(spread * LANGUAGE_WEIGHT)

model.Minimize(sum(penalties))

solver = cp_model.CpSolver()
# Optional: Set a time limit so it doesn't run forever on huge datasets
solver.parameters.max_time_in_seconds = 30.0 

status = solver.Solve(model)

if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
    lines = []
    lines.append(f"Solution found (Status: {solver.StatusName(status)})")
    lines.append("")
    
    # Organize results by class for better readability
    class_grouping = {c: [] for c in classes}
    
    for s in students:
        for c in classes:
            # Check if the solver set this assignment to 1 (True)
            if solver.Value(assignments[(s["id"], c)]) == 1:
                class_grouping[c].append(s)

    # Build the report
    for c in classes:
        lines.append(f"=== {c} ===")
        lines.append(f"{'Name':<10} | {'Sex':<5} | {'Origin':<10} | {'Subject':<10} | {'Language':<18}")
        lines.append("-" * 70)
        for student in class_grouping[c]:
            lines.append(
                f"{student['name']:<10} | {student['sex']:<5} | {student['origin']:<10} | "
                f"{student['subject']:<10} | {student['language']:<18}"
            )
        lines.append(f"Total: {len(class_grouping[c])} students")
        lines.append("")
        
    # Print the total 'penalty' score if you have soft rules
    lines.append(f"Total Penalty Score: {solver.ObjectiveValue()}")

    report = "\n".join(lines)
    print(report)

    # Also write results to a .txt file
    cfg.OUTPUT_TXT_PATH.write_text(report + "\n", encoding="utf-8")
    print(f"\nWrote results to: {cfg.OUTPUT_TXT_PATH}")

    # Write results to a .csv file (class name in first column)
    with cfg.OUTPUT_CSV_PATH.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["class", "name", "sex", "origin", "subject", "language"])
        for c in classes:
            for student in class_grouping[c]:
                writer.writerow([
                    c,
                    student["name"],
                    student["sex"],
                    student["origin"],
                    student["subject"],
                    student["language"],
                ])
    print(f"Wrote CSV results to: {cfg.OUTPUT_CSV_PATH}")

else:
    print("Could not find a solution that satisfies all hard constraints.")