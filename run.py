from MyApp import app
from waitress import serve

if __name__ == '__main__':
    # serve(app,host='0.0.0.0',port=50100,threads=2)
    #app.run(host='10.120.56.105',debug=True)
    app.run(host='0.0.0.0',debug=True)
