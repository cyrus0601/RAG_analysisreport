"""
Description: vectorize the extracted content from the PDF file and store it in a database for future use.
"""
# %%
import os
import pdfplumber

from openai import AzureOpenAI
from langchain_openai import AzureOpenAIEmbeddings
from langchain.vectorstores import Chroma
from langchain.docstore.document import Document
from chromadb.config import Settings

os.environ["AZURE_OPENAI_API_KEY"] = ""
os.environ["AZURE_OPENAI_API_VERSION"] = ""
os.environ["AZURE_OPENAI_ENDPOINT"] = ""
azure_openai_embeddings = AzureOpenAIEmbeddings(
        deployment='text-embedding-3-large',
        model="text-embedding-3-large",
        azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
        openai_api_type=os.getenv("OPENAI_API_TYPE"),
        openai_api_key=os.getenv("AZURE_OPENAI_API_KEY"),
        chunk_size=1024,
    )

# %%
def extract_content(input_pdf, pages_to_extract):
    content = []

    with pdfplumber.open(input_pdf) as pdf:
        for page_num in pages_to_extract:
            page = pdf.pages[page_num]
            page_content = ""
            text = page.extract_text()

            if text:
                page_content += text.strip() + "\n"

            tables = page.extract_tables()

            for table in tables:
                if table:
                    table_str = "\n".join(
                        ["\t".join([cell if cell is not None else "-" for cell in row]) for row in table if row]
                    )
                    page_content += f"--------------------------------------------------\nTable:\n{table_str}\n--------------------------------------------------\n"

            content.append(page_content)

    # Combine every two pages into one unit
    combined_content = []
    for i in range(0, len(content) - 1):
        combined_content.append(content[i] + content[i + 1])

    return combined_content

def vector_db(extracted_content, azure_openai_embeddings, persist_directory):
    documents = [Document(page_content = content) for content in extracted_content]
    persist_directory = persist_directory 

    vector_store = Chroma.from_documents(
        documents=documents,
        embedding=azure_openai_embeddings,
        persist_directory=persist_directory,
    )
    vector_store.persist()

# %%
input_pdf = ".pdf"
realtek_pages_to_extract = list(range(67+3, 89+3))
realtek_extracted_content = extract_content(input_pdf, realtek_pages_to_extract)
vector_db(realtek_extracted_content, azure_openai_embeddings, "db/")

input_pdf = ".pdf"
airhona_pages_to_extract = list(range(47+4, 57+4))  
airhona_extracted_content = extract_content(input_pdf, airhona_pages_to_extract)
vector_db(airhona_extracted_content, azure_openai_embeddings, "db/")

