from services.equipment_report_service import (
    get_enabled_department_summary,
    get_enabled_equipment_summary,
)


def get_dashboard_report(session):
    summaries = get_enabled_equipment_summary(session)
    return {
        "equipment": summaries,
        "departments": get_enabled_department_summary(session),
    }
