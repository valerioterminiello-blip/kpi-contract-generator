# Agreement date
if agreement_date:
    try:
        dt = datetime.strptime(agreement_date, "%Y-%m-%d")  # CHANGED
        formatted_agreement_date = dt.strftime("%d %B %Y")
    except ValueError:
        formatted_agreement_date = agreement_date

# Start date
try:
    dt = datetime.strptime(start_date, "%Y-%m-%d")  # CHANGED
    formatted_start = dt.strftime("%d %B %Y")
except ValueError:
    formatted_start = start_date

# End date
try:
    dt_end = datetime.strptime(end_date, "%Y-%m-%d")  # CHANGED
    formatted_end = dt_end.strftime("%d %B %Y")
except ValueError:
    formatted_end = end_date if end_date else "To be agreed"
