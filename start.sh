#!/bin/bash

if [ ! -f instance/.initialized ]; then
  mkdir -p instance
  python create_admin.py
  touch instance/.initialized
fi

gunicorn app:app
