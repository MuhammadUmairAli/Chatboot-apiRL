import os
from langchain_community.document_loaders import UnstructuredFileLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS
from langchain_openai import OpenAIEmbeddings, ChatOpenAI
from langchain.chains import RetrievalQA
from dotenv import load_dotenv
from os

load_dotenv()
# Get the key from environment variable
api_key = os.getenv("OPENAI_API_KEY")

if not api_key:
    raise ValueError("OPENAI_API_KEY is not set in environment.")

# Load & split document
# print("Starting to load document...")  # Debug
loader = UnstructuredFileLoader("Ruby Leather.docx")
docs = loader.load()
# print("Document loaded.")

splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
chunks = splitter.split_documents(docs)

# Create vector DB with OpenAI embeddings
embeddings = OpenAIEmbeddings()
vector_db = FAISS.from_documents(chunks, embeddings)

# Build retriever + chain
retriever = vector_db.as_retriever()
qa_chain = RetrievalQA.from_chain_type(
    llm=ChatOpenAI(model_name="gpt-4o"),  # GPT-4.1, cheaper & better
    retriever=retriever
)


# Expose a function to answer questions
def get_response(prompt: str) -> str:
    return qa_chain.run(prompt)
# Ask a sample question
# query = "I want to know about Accessories"
# response = qa_chain.run(query)
# print("Answer:", response)
