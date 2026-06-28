from datetime import datetime, timedelta, timezone

from fastapi.templating import Jinja2Templates


BEIJING_TZ = timezone(timedelta(hours=8))


def format_beijing_time(value):
    if not value:
        return ""
    if isinstance(value, str):
        try:
            value = datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            return value
    if not isinstance(value, datetime):
        return value
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(BEIJING_TZ).strftime("%Y-%m-%d %H:%M:%S")


def create_templates() -> Jinja2Templates:
    templates = Jinja2Templates(directory="app/web/templates")
    templates.env.filters["localtime"] = format_beijing_time
    return templates
