#!/usr/bin/env bash
set -o errexit

pip install -r requirements.txt

DJANGO_SETTINGS_MODULE=excelsior_system.settings python manage.py collectstatic --no-input
DJANGO_SETTINGS_MODULE=excelsior_system.settings python manage.py migrate