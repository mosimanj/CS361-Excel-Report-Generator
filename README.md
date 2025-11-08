# Excel Report Generator
The Excel Report Generator (ERG) is a microservice that creates and formats an Excel report based on received data. 
Clients can request an Excel report by specifying headers, rows (data), whether the report should be sorted by a column,
and whether one of the styling templates should be applied. The ERG responds with the absolute file path of the
generated report. 

## Communication Contract
The ERG utilizes ZeroMQ to handle communication between itself and the client. Continue reading below for instructions 
on requesting and receiving data, along with example calls. 

## Requesting Data

### Request Instructions
1. Import the zmq and json packages. 
2. Open a new request socket connected to localhost port number 5727.
3. Prepare your data in the required format (see _Request Format_ section below). 
4. Convert the dictionary to JSON format. 
5. Send the JSON content through your connected socket. 

### Request Format
```javascript
example_request = {
    'Headers': ['header1', 'header2', 'header3', 'header4'], 
    'Rows': [
        ['row1-col1', 'row1-col2', 'row1-col3', 'row1-col4'], 
        ['row2-col1', 'row2-col2', 'row2-col3', 'row2-col4']
        // Etc.
    ],
    'sort_by': 'header2', // must match the name of one of the Header values.
    'style': 'templateName' // must match the name of one of the pre-defined templates. 
}
```
_Request Format Notes:_

- If the Excel file doesn't need to be sorted by a column or doesn't need a style template applied, omit those keys entirely.
- Each row must have the same number of entries as the number of headers (i.e. each row must have a value for each header item).
- If a header is provided for sort_by, the rows will be sorted by their value in ascending order. 

### Example Call (Request)
```python
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
    'style': 'Blue'
}

# Convert request to JSON & send
request_json = json.dumps(request)
socket.send_string(request_json)
```

## Receiving Data

### Response Instructions

[//]: # (TODO: Complete)

### Example Call (Response)
```python

```
[//]: # (TODO: Complete)

## UML Sequence Diagram

[//]: # (TODO: Complete)

## Style Templates
Below is an overview of the name of each template, along with a preview of their appearance.

[//]: # (TODO: Add screenshots of available templates)