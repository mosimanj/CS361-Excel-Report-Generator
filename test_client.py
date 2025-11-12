import zmq
import json

# Setup communication (client)
context = zmq.Context()
socket = context.socket(zmq.REQ)
socket.connect("tcp://localhost:5727")

# Example report request
request = {
    'Headers': ['ID', 'Date', 'Name', 'Sale Amount'],
    'Rows': [
        [7, '2024/10/05', 'John Doe', 35.24],
        [4, '2024/07/02', 'Jim Doe', 70],
        [10, '2024/12/15', 'Jane Doe', 25]
    ],
    'sort_by': 'ID',
    'style': 'blue'
}

# Convert request to JSON & send
request_json = json.dumps(request)
print(f'Client sending request: {request_json}')
socket.send_string(request_json)

# Receive filepath and decode back to a string
message = socket.recv().decode('utf-8')
print(f'Client received message: {message}')