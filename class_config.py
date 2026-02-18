"""
Central configuration for `class-optimisation.py`.

Edit values here to tune the model without changing solver logic.
"""

from pathlib import Path

# Input
CSV_FILENAME = "Klassedannelse_Kristian.csv"
CSV_PATH = Path(__file__).with_name(CSV_FILENAME)

# Output
OUTPUT_TXT_FILENAME = "class_optimisation_results.txt"
OUTPUT_TXT_PATH = Path(__file__).with_name(OUTPUT_TXT_FILENAME)
OUTPUT_CSV_FILENAME = "class_optimisation_results.csv"
OUTPUT_CSV_PATH = Path(__file__).with_name(OUTPUT_CSV_FILENAME)

# Classes
CLASSES = [
    "Class_A",
    "Class_B",
    "Class_C",
    "Class_D",
    "Class_E",
    "Class_F",
    "Class_G",
    "Class_H",
    "Class_I",
]

# Hard constraints
MAX_CLASS_SIZE = 32
MIN_CLASS_SIZE = 20
MIN_FEMALE_PER_CLASS = 3
MIN_MALE_PER_CLASS = 3

# Soft constraints (objective weights)

# the weight for the otpmisation of ORIGIN_MAX_PER_CLASS
ORIGIN_WEIGHT = 5
# the weight associated with groupng students with the same subject
SUBJECT_WEIGHT = 5
# Prefer grouping same language
LANGUAGE_WEIGHT = 5
# Balance class sizes around the average (number of students / number of classes )
CLASS_BALANCE_WEIGHT = 2

# Soft constraint parameters
# Prefer at most this many students with the same origin in a class.
ORIGIN_MAX_PER_CLASS = 2

