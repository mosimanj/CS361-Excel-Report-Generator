import json
import zmq
from erg import ReportGenerator

#Setup communication (server)
context = zmq.Context()
socket = context.socket(zmq.REP)
socket.bind("tcp://*:5727")

while True:
    # Receive message
    message = socket.recv_string()
    request = json.loads(message)

    # Create report generator & perform operations
    generator = ReportGenerator(request)

    if generator.needs_sort():
        generator.sort_report()

    if generator.needs_style():
        generator.style_report()

    file_path = generator.generate_excel()

    # Send filepath of generated Excel file to client
    socket.send_string(file_path)