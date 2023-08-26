from flask import Flask

# initialize flask
def create_app():
    app = Flask(__name__)
    # initialize your app with a secret key
    app.config['SECRET_KEY'] = 'Loxahatchee15846'

    return app