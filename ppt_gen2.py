import streamlit as st
from ppt_lib import *
import base64
import re
import openai

def create_download_link(data, filename):
    b64 = base64.b64encode(data).decode()  # 將檔案數據轉換為 base64 編碼
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{filename}">點此下載</a>'
    return href

def split_list(input_list, chunk_size):
    return [input_list[i:i + chunk_size] for i in range(0, len(input_list), chunk_size)]
    
def remove_punctuation(text):
    # 使用正則表達式移除標點符號
    cleaned_text = re.sub(r'[^\w\s]', '', text)
    return cleaned_text

if __name__ == '__main__':
    if 'context' not in st.session_state:
        st.session_state.context = ''

    st.set_page_config(layout = 'wide')
    
    st.title("ChatPPT")
    
    col6, col7 = st.columns(2)
    
    title = col6.text_input("請輸入標題:", value = '要如何學習機器學習')
    sub_title = col7.text_input("請輸入副標題:", value = '李御璽/銘傳大學資工系')
    
    col8, col9 = st.columns([1,4])
    
    context_title = col8.text_input("請輸入內文標題:", value = '學習機器學習的步驟')
    context = col9.text_area("請輸入內文:", value = st.session_state.context, height=170)
    
    col1, col2, col3, col4, col5 = st.columns(5)
    chatGPT = col1.checkbox('ChatGPT導入內文', value=False)
    topic = col2.text_input("您希望ChatGPT是何種專家:", value = '機器學習')
    temperature = col3.slider('ChatGPT的創意程度', 0.0, 1.0, 0.0, 0.1)
    num_of_points = col5.number_input("投影片每頁要列出幾點:", value = 4)
    
    if st.button("產生PPT"):
        if chatGPT:
            # 呼叫 API 
            if len(topic) >= 1 and ord(topic[0]) >= 256: # 是否為中文
                system_prompt = f"你是{topic}專家"
            else:
                system_prompt = f"You are a {topic} assistant"
                
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": title}
            ]
            
            response = openai.ChatCompletion.create(
                model = 'gpt-3.5-turbo',
                messages = messages,
                temperature = temperature,
            )
            
            context = response.choices[0].message.content
            st.session_state.context = context
            
            col4.checkbox('ChatGPT導入顯示')
    
        # 產生投影片
        prs = Presentation()
        
        sub_title = sub_title.replace('/', '\n')
        
        create_title(prs, title, sub_title)
        
        if len(context) > 0:
            item_list = []
            
            for line in context.splitlines():
                item_list.append(line.strip())
            
            chunked_lists = split_list(item_list, num_of_points)

            slides = []

            # Print the chunked lists
            for sublist in chunked_lists:
                slides.append({'title': context_title, 'content': sublist})
            
            create_body(prs, slides)
    
            filename = remove_punctuation(title) + '.pptx'
            prs.save(filename)   

            with open(filename, 'rb') as f:
                data = f.read()

            #顯示下載連結
            st.markdown(create_download_link(data, filename), unsafe_allow_html=True)