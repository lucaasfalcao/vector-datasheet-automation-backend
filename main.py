"""Aplication initializing function and printing"""

import uvicorn
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from controllers import routes
from config.middlewares import middlewares, cors, log_request_middleware


DESCRIPTION = """
This is a Automation service dedicated to Vector FIDC

Developed by Lucas Falcão.
"""

app = FastAPI(
    title="Vector FIDC Automation Service",
    description=DESCRIPTION,
    version='0.0.1',
    contact={
        'name': 'Lucas Falcão',
        'email': 'suporte@vectorfidc.com.br',
    },
    swagger_ui_parameters={'defaultModelsExpandDepth': -1},
)
for middleware in middlewares:
    app.middleware(middleware.type)(middleware(app))
app.add_middleware(CORSMiddleware, **cors)
app.middleware('http')(log_request_middleware)

for router in routes:
    app.include_router(router)

print("""
Service is available on port 8000:

  Docs (Swagger) available on: http://localhost:8000/docs.
""")

if __name__ == '__main__':
    uvicorn.run(
        'main:app',
        host='0.0.0.0',
        port=8000,
        reload=True,
    )
