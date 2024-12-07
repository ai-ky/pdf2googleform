import fitz  # PyMuPDF
import re
import os
import io
from PIL import Image
import base64

# 打開PDF檔案
pdf_path = 'scienceQuiz.pdf'  # 請將此路徑替換為您的PDF檔案路徑
doc = fitz.open(pdf_path)

# 建立保存圖片的資料夾
image_folder = 'images'
if not os.path.exists(image_folder):
    os.makedirs(image_folder)

question_number = 1
text_buffer = ''
image_counter = 1
image_info_list = []

# 存儲所有題目信息的列表
questions = []

for page_index in range(len(doc)):
    page = doc[page_index]
   
    page_dict = page.get_text("dict")
    blocks = page_dict["blocks"]
    for b_index, b in enumerate(blocks):
        print(b)
        block_type = b["type"]
        if block_type == 0:
            text = ''
            for line in b["lines"]:
                for span in line["spans"]:
                    text += span["text"]
                text += '\n'
            lines = text.strip().split('\n')
            
            for line in lines:
                if re.match(r'^\d+\.', line.strip()):
                    if text_buffer:
                        question_data = {}
                        question_data['number'] = question_number
                        question_data['text'] = text_buffer.strip()
                        question_data['images'] = image_info_list.copy()
                        questions.append(question_data)
                        question_number += 1
                        text_buffer = ''
                        image_info_list = []
                    text_buffer += line.strip() + '\n'
                else:
                    text_buffer += line.strip() + '\n'

        elif block_type == 1:  # 圖片塊
            # 嘗試獲取圖片的 xref 編號
            xref = b.get("xref")
            if xref is None:
                # 如果沒有 xref，嘗試從 src 中解析
                src = b.get("src")
                if src and "xref" in src:
                    xref = int(src.split(" ")[-1])
                else:
                    # 無法獲取 xref，嘗試直接從塊中獲取圖片資料
                    image_data = b.get("image")
                    if image_data:
                        image_bytes = image_data
                        image_ext = 'png'  # 預設使用 png 格式
                    else:
                        # 無法獲取圖片資料，跳過此圖片塊
                        continue
            else:
                # 使用 xref 提取圖片
                base_image = doc.extract_image(xref)
                image_bytes = base_image['image']
                image_ext = base_image['ext']

            # 保存圖片
            try:
                image = Image.open(io.BytesIO(image_bytes))
                image_filename = f"{image_counter}.{image_ext}"
                image_path = os.path.join(image_folder, image_filename)
                image.save(image_path)
                image_counter += 1

                bbox = b["bbox"]
                position = bbox  # (x0, y0, x1, y1)
                image_info_list.append({
                    'filename': image_filename,
                    'position': position,
                    'base64': base64.b64encode(image_bytes).decode('utf-8'),
                    'ext': image_ext
                })

            except Exception as e:
                print(f"提取圖片失敗：{e}")

    # 檢查是否是最後一個塊，如果是，且有未輸出的題目，則輸出
    if page_index == len(doc) - 1:
        if text_buffer:
            # 處理並保存最後一題
            question_data = {}
            question_data['number'] = question_number
            question_data['text'] = text_buffer.strip()
            question_data['images'] = image_info_list.copy()
            questions.append(question_data)
            text_buffer = ''
            image_info_list = []

# 生成 Google Apps Script 程式碼
script_lines = []
script_lines.append("// 這是自動生成的Google Apps Script程式碼，用於建立表單")
script_lines.append("function createForm() {")
script_lines.append("  var form = FormApp.create('自動生成的表單');")
script_lines.append("  form.setIsQuiz(true);  // 將表單設為測驗模式")

# 添加姓名和學號兩個簡答題
script_lines.append("  // 姓名")
script_lines.append("  var nameItem = form.addTextItem();")
script_lines.append("  nameItem.setTitle('姓名').setRequired(true);")
script_lines.append("  // 學號")
script_lines.append("  var idItem = form.addTextItem();")
script_lines.append("  idItem.setTitle('學號').setRequired(true);")

for q in questions:
    script_lines.append(f"  // 題目 {q['number']}")
    # 處理題目描述，將換行符替換為 '\\n' 並進行轉義
    question_text = q['text'].replace('\n', '\\n').replace("'", "\\'")
    # 創建單選題
    script_lines.append("  var item = form.addMultipleChoiceItem();")
    script_lines.append(f"  item.setTitle('{question_text}').setRequired(true);")
    script_lines.append("  item.setChoices([")
    script_lines.append("    item.createChoice('A'),")
    script_lines.append("    item.createChoice('B'),")
    script_lines.append("    item.createChoice('C'),")
    script_lines.append("    item.createChoice('D')")
    script_lines.append("  ]);")
    if q['images']:
        # 創建圖片項並添加圖片
        for img_info in q['images']:
            img_data = img_info['base64']
            img_ext = img_info['ext']
            # 創建圖片 Blob
            script_lines.append("  var imgBlob = Utilities.newBlob(Utilities.base64Decode('{}'), 'image/{}', 'image{}');".format(
                img_data, img_ext, img_info['filename']))
            # 添加圖片項
            script_lines.append("  var imgItem = form.addImageItem().setImage(imgBlob);")
            
    # 設置正確答案（可選，如果需要設置答案，可在此處修改）
    # script_lines.append("  item.setCorrectAnswer('A');")  # 預設將 A 設為正確答案

script_lines.append("}")

# 將程式碼內容輸出到檔案
with open('script.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(script_lines))

print("已生成Google Apps Script程式碼，保存在script.txt中。")
