"""
Description: This script is used to generate a report based on the given information.
"""
# %%
import os
import pandas as pd

from openai import AzureOpenAI
from langchain_openai import AzureOpenAIEmbeddings
from langchain.vectorstores import Chroma
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
def load_vector_db(persist_directory, azure_openai_embeddings):
    db = Chroma(persist_directory=persist_directory, embedding_function=azure_openai_embeddings)

    return db

def retrieval(db, query):
    results_with_scores = db.similarity_search_with_score(query, k=1)
    threshold = 0.7
    matching_docs = [(doc, score) for doc, score in results_with_scores if score > threshold]

    def print_documents_with_index(documents):
        info = ""

        for index, doc in enumerate(documents, start=1):
            info += f"[參考文件{index}]:\n\n{doc[0].page_content}\n"

        return info

    return print_documents_with_index(matching_docs)

def ask_question(db, template, mission=None):
        result = ""
        message_text = []
        

        if template=="summary":
            message_text = [
                {
                    "role": "system",
                    "content":"""你是一個分析公司商業分析報告的AI助手。透過找出各個分析資訊的關聯來產出這家公司的優勢和潛在風險，將其整齊列出並說明一名投資者想要看到的對公司的優勢劣勢和發展趨勢。"""
                },
                {
                    "role": "user",
                    "content": f"以下是提供的商業分析報告內容\n\n{mission}\n\nAssistant:"
                }
            ]
        elif template=="analysis":
            info = retrieval(db, mission)
            message_text = [
                {
                    "role": "system",
                    "content": """你是一個基於提供的資訊進行整合，然後生成商業分析報告的繁體中文AI助手，回答過程只能使用提供給你的資料並且使用時要確保資料的使用是正確的。若你認為提供資料中的表格和你要分析的問題相關，那幫我用plotly繪圖套件生成圖表，輸出可以直接執行的程式碼，記得import所有要用到的套件。並且固定輸出以下格式方便我後續使用:\n根據提供的資訊我的分析是：\n\n{會呈現出來的分析文字}\n參考頁數:{頁數}\n\n=====\n\n{code內容}(如果有要畫圖表)。如果沒有要畫圖表則code內容直接空白，提供資料沒有任務答案直接在分析只說"沒有相關知識"，因為有兩部分無論有沒有內容中間要用=====分開來，記得不要產生不存在於提供資料的知識內容和胡言亂語，數值部分的逗點在畫表時記得移除。
                    """
                },
                {
                    "role": "user",
                    "content": f"任務:{mission}，以下是提供的資訊 :\n{info}\n\nAssistant:"
                }
            ]
        else:
            return "Invalid template type"

        client = AzureOpenAI(
            azure_endpoint = "", 
            api_key=os.environ[""],  
            api_version=""
        )
        
        response = client.chat.completions.create(
            model="gpt4-o", 
            messages=message_text,
            stream=True,
        )

        for chunk in response:

            if chunk.choices:
                first_choice = chunk.choices[0]

                if first_choice.delta and first_choice.delta.content is not None:
                    result += first_choice.delta.content
        
        return result

def process_analysis(x_return_list, path):
    prefix = "根據提供的資訊我的分析是：\n\n"
    analysis_list = []
    code_list = []

    for item in x_return_list:
        parts = item.split('=====')
        analysis = parts[0].strip()

        if analysis.startswith(prefix):
            analysis = analysis[len(prefix):]  
        analysis_list.append(analysis)

        if len(parts) > 1:
            code = parts[1].strip()
            code_lines = code.split('\n')
            code_lines = [line for line in code_lines if "```" not in line.strip()]
            cleaned_code = "\n".join(code_lines)
            code_list.append(cleaned_code)
        else:
            code_list.append('')

    df = pd.DataFrame({
        'Analysis': analysis_list,
        'Code': code_list
    })

    df.to_csv(path, index=False, encoding='utf-8-sig')

    return analysis_list

def concat_summary(summary, path):
    df = pd.read_csv(path, encoding="utf-8-sig")
    new_row = pd.DataFrame({"Analysis": [summary], "Code": [""]})
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_csv(path, index=False, encoding="utf-8-sig")


#================================================================================================


# %%
x_db = load_vector_db("db/x_db", azure_openai_embeddings)
y_db = load_vector_db("db/y_db", azure_openai_embeddings)
x_return_list = []
y_return_list = []

query_queue = [ "公司的業務之主要內容說明。", "公司主要產品的重要用途是什麼？",  "公司的主要銷售地區在哪裡?", "公司近兩年的銷售量及銷售值是多少？", "公司預測市場未來的供需及成長性如何？", "公司計畫開發的新產品有什麼?詳細列出不要遺漏", "公司短期和長期的業務發展計畫是什麼？", "公司從業員工及學歷情況如何？"]

for query in query_queue:
    result = ask_question(x_db, "analysis", query)
    x_return_list.append(result)

for query in query_queue:
    result = ask_question(y_db, "analysis", query)
    y_return_list.append(result)


# %%
x_analysist_list = process_analysis(x_return_list, "x_analysis.csv")
summary = ask_question(x_db, "summary", x_analysist_list)
concat_summary(summary, "x_analysis.csv")

# %%
y_analysist_list = process_analysis(y_return_list, "y_analysis.csv")
summary = ask_question(y_db, "summary", y_analysist_list)
concat_summary(summary, "y_analysis.csv")

# %%