#!/usr/bin/env python3
"""Test export of deck 8 using the updated logic"""

import sys
import os

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import export_deck
from flask import Flask

app = Flask(__name__)

with app.app_context():
    # Simulate a Flask request context for format parameter
    from flask import request
    from werkzeug.test import EnvironBuilder
    
    builder = EnvironBuilder(path='/api/export_deck/7?format=pptx')
    env = builder.get_environ()
    
    with app.request_context(env):
        result = export_deck(7)
        print(result)
