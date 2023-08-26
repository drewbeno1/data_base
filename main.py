from Website import create_app

app = create_app()

if __name__ == '__main__':
    # turn this off when we are in Production 
    app.run(debug=True)