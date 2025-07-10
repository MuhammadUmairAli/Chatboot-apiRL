from fastapi import FastAPI
from pydantic import BaseModel
from chatbot import get_response
from fastapi.responses import FileResponse

app = FastAPI()

class PromptRequest(BaseModel):
    query: str

@app.post("/ask")
async def ask_question(request: PromptRequest):
    print("Received:", request.query)
    response = get_response(request.query)
    return {"response": response}

@app.get("/")
def read_index():
    return FileResponse("index.html")
