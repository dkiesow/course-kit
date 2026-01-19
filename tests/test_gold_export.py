#!/usr/bin/env python3
"""Test PPTX export for deck with gold slides."""

import sys
sys.path.insert(0, '.')

from app import app

with app.app_context():
    with app.test_request_context('/?format=pptx'):
        # Export deck 11 which has gold slides
        from flask import request
        from app import export_deck
        response = export_deck(11)
        print(response)
