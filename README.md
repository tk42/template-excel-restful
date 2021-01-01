# template-excel-restful
template for porting excel VBA to RESTful API

## :bulb: Basic idea
Everyone loves Excel, but VBA is disliked. In particular, VBA is difficult to debug and manage code.
So I came up with the idea of preparing a backend server and using Excel with custom functions as the frontend.
In this Excel, REST-API is only executed from VBA via HTTP, the actual calculation is executed on the backend server, and only the result is returned to Excel again.

Also, if you are using Office 365 or later, you can use JavaScript in addition to VBA, so you can display the results in real time by Websocket communication.

## Policy
Build FastAPI or Flask, a micro web framework written in Python, and send and receive messages converted to JSON format by VBA.
Supported not only REST-API but also WebSocket

## Requiements
### Python
 - FastAPI
 https://fastapi.tiangolo.com/ja/
 - flask-restx
 https://flask-restx.readthedocs.io/en/latest/
 - websockets
 https://websockets.readthedocs.io/en/stable/intro.html
 
### Excel
 - VBA-JSON
 
 https://github.com/VBA-tools/VBA-JSON
 - "Excel でカスタム関数を作成する"
 
 https://support.microsoft.com/ja-jp/office/excel-%E3%81%A7%E3%82%AB%E3%82%B9%E3%82%BF%E3%83%A0%E9%96%A2%E6%95%B0%E3%82%92%E4%BD%9C%E6%88%90%E3%81%99%E3%82%8B-2f06c10b-3622-40d6-a1b2-b6748ae8231f
 - "Excel JavaScript API を使用した基本的なプログラミングの概念"
 
 https://docs.microsoft.com/ja-jp/office/dev/add-ins/excel/excel-add-ins-core-concepts
 - "カスタム関数でデータを受信して処理する" (by JavaScript)
 
 https://docs.microsoft.com/ja-jp/office/dev/add-ins/excel/custom-functions-web-reqs
