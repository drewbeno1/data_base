from flask import Flask

# initialize flask
def create_app():
    app = Flask(__name__)
    # initialize your app with a secret key
    app.config['SECRET_KEY'] = 'Loxahatchee15846'

    # register our blueprints (from views and auth)
    from .views import views
    from .auth import auth

    # just leave it as slash so there is no prefix for the functions
    app.register_blueprint(views, url_prefix='/')
    app.register_blueprint(auth, url_prefix='/')

    return app