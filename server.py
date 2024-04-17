from waitress import serve
from app import app 
print("Server started on port 8040")
serve(app, host='0.0.0.0', port=8040)
