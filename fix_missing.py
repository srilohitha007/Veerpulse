import os

app_file = r'c:\Users\SRI LOHITHA\OneDrive\Desktop\veerpulse\app\__init__.py'

with open(app_file, 'r', encoding='utf-8') as f:
    content = f.read()

replacements = [
    (
"""                current_training_hours = round(
                    sum(
                        record.training_hours or 0
                        for record in current_records
                    ),
                    1
                )""",
"""                _t_vals = [r.training_hours for r in current_records if r.training_hours is not None]
                current_training_hours = round(sum(_t_vals), 1) if _t_vals else None"""
    ),
    (
"""        total_duty_hours = round(
            sum(
                item.duty_hours or 0
                for item in operational_records
            ),
            1
        )""",
"""        _d_vals = [i.duty_hours for i in operational_records if i.duty_hours is not None]
        total_duty_hours = round(sum(_d_vals), 1) if _d_vals else None"""
    ),
    (
"""        total_training_hours = round(
            sum(
                item.training_hours or 0
                for item in operational_records
            ),
            1
        )""",
"""        _tr_vals = [i.training_hours for i in operational_records if i.training_hours is not None]
        total_training_hours = round(sum(_tr_vals), 1) if _tr_vals else None"""
    ),
    (
"""        current_total_duty = round(
            sum(
                item.duty_hours or 0
                for item in current_records
            ),
            1
        )""",
"""        _c_d_vals = [i.duty_hours for i in current_records if i.duty_hours is not None]
        current_total_duty = round(sum(_c_d_vals), 1) if _c_d_vals else None"""
    ),
    (
"""        current_training_hours = round(
            sum(
                item.training_hours or 0
                for item in current_records
            ),
            1
        )""",
"""        _c_t_vals = [i.training_hours for i in current_records if i.training_hours is not None]
        current_training_hours = round(sum(_c_t_vals), 1) if _c_t_vals else None"""
    )
]

for old, new in replacements:
    content = content.replace(old, new)

with open(app_file, 'w', encoding='utf-8') as f:
    f.write(content)

print("Done replacements.")
