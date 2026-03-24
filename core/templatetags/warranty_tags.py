"""
core/templatetags/warranty_tags.py

Usage in templates:
    {% load warranty_tags %}
    {% warranty_status r.date_due as ws %}
    {% if ws %}
      <span class="warranty-badge warranty-{{ ws.level }}">{{ ws.label }}</span>
    {% endif %}
"""

from django import template
from django.utils import timezone
import datetime

register = template.Library()


@register.simple_tag
def warranty_status(date_due):
    """
    Returns a dict with level and label, or None if no date / not near expiry.

    Levels:
        'critical'  — due today or already overdue
        'urgent'    — within 7 days
        'warning'   — within 30 days
    """
    if not date_due:
        return None

    today = timezone.localdate()

    # Accept both date and datetime objects
    if isinstance(date_due, datetime.datetime):
        due = date_due.date()
    else:
        due = date_due

    delta = (due - today).days

    if delta <= 0:
        return {'level': 'critical', 'label': '🔴 Due Today' if delta == 0 else '🔴 Due', 'days': delta}
    elif delta <= 7:
        return {'level': 'urgent', 'label': f'🟠 Due in {delta}d', 'days': delta}
    elif delta <= 30:
        return {'level': 'warning', 'label': f'🟡 Due in {delta}d', 'days': delta}

    return None