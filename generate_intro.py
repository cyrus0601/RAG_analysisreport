"""
Description: This script is used to extract the information between two keywords from a PDF file 
and then use the extracted information to generate a summary of the content using the Azure OpenAI API. The summary is then written to a text file.
"""
# %%
import pdfplumber
import os

from openai import AzureOpenAI

# %%
os.environ["AZURE_OPENAI_API_KEY"] = ""
os.environ["AZURE_OPENAI_API_VERSION"] = ""
os.environ["AZURE_OPENAI_ENDPOINT"] = ""

print("AZURE_OPENAI_API_KEY:", os.getenv("AZURE_OPENAI_API_KEY"))
print("AZURE_OPENAI_API_VERSION:", os.getenv("AZURE_OPENAI_API_VERSION"))
print("AZURE_OPENAI_ENDPOINT:", os.getenv("AZURE_OPENAI_ENDPOINT"))

# %%
def extract_between_keywords(pdf_path:str, start_word:str, end_word:str)->str:
    text_between_keywords = ""
    whole_text = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:15]:
            whole_text += page.extract_text()

        first_start_idx = whole_text.find(start_word)
        first_end_idx = whole_text.find(end_word, first_start_idx + len(start_word))

        second_start_idx = whole_text.find(start_word, first_end_idx + len(end_word))
        second_end_idx = whole_text.find(end_word, second_start_idx + len(start_word))

        if second_start_idx != -1 and second_end_idx != -1:
            text_between_keywords = whole_text[second_start_idx + len(start_word):second_end_idx].strip()
    
    return text_between_keywords

def execute_LLM(info:str, mission:str)->str:
    message_text = [
        {
            "role": "system",
            "content": "你是一個基於提供的資訊進行整合，然後協助商業分析報告者對公司發展進行分析的繁體中文AI助手。對於生成的回答大標題用\'一、二、三...表示\'，列點的部分用\'-\'當開頭"
        },
        {
            "role": "user",
            "content": f"{mission}，以下是提供的資訊 : {info}\n\nAssistant:"
        }
    ]
    
    client = AzureOpenAI(
      azure_endpoint = "", 
      api_key=os.environ[""],  
      api_version=""
    )

    response = client.chat.completions.create(
        model="gpt4-3", 
        messages=message_text,
        stream=True,
    )
        
    result = ""
    for chunk in response:
        if chunk.choices:
            first_choice = chunk.choices[0]
            if first_choice.delta and first_choice.delta.content is not None:
                print(first_choice.delta.content, end="", flush=True)
                result += first_choice.delta.content
    
    return result

def write_summary_2_txt(path_2_txt:str, content:str):
    path_2_txt = path_2_txt
    with open(path_2_txt, 'w', encoding='utf-8') as file:
        file.write(content)


#====================================================================================================


# %%
info = extract_between_keywords("", "致股東報告書", "公司簡介")
content = execute_LLM(info, "年度報告，幫我把提供的內容做整理並提高可讀性，過程中不要遺漏任何重要資訊，也不要生成任何錯誤或不存在的內容，若必要的話你也可以用列點的方式完成任務，並且關於產品、技術、產業趨勢分析的相關內容請詳細列出不要省略")
write_summary_2_txt("", content)

# %%
info = extract_between_keywords("", "致股東報告書", "公司簡介")
content = execute_LLM(info, "年度報告，幫我把提供的內容做整理並提高可讀性，過程中不要遺漏任何重要資訊，也不要生成任何錯誤或不存在的內容，若必要的話你也可以用列點的方式完成任務，並且關於產品、技術、產業趨勢分析的相關內容請詳細列出不要省略")
write_summary_2_txt("", content)
