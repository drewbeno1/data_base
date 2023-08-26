# Create Routes - where users can actually go to (Home Page)
from flask import Blueprint

# this is a blueprint for our first page of our flask web app
views = Blueprint('views', __name__)

@views.route('/')
# When this path is navigated to, this function will run
def home():
    return "<h1>Test</h1>"