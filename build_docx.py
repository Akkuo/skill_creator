import os
import re
import sys
import base64
import json
import urllib.request
import subprocess

def create_api_url(text):
    data = {"code": text, "mermaid": {"theme": "base", "look": "handDrawn"}}
    json_str = json.dumps(data)
    encoded = base64.urlsafe_b64encode(json_str.encode('utf-8')).decode('ascii').replace("=", "")
    return f"https://mermaid.ink/img/{encoded}"

def build(md_file_path):
    if not os.path.exists(md_file_path):
        print(f"找不到檔案：{md_file_path}", flush=True)
        return

    # 計算路徑，保證腳本可以在任意地點呼叫
    work_dir = os.path.dirname(os.path.abspath(md_file_path)) or '.'
    base_name = os.path.splitext(os.path.basename(md_file_path))[0]

    with open(md_file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 移除「最終輸出與自動化交接 (Export Protocol)」章節，不輸出到 Word
    content = re.sub(
        r'##\s+\S*[\d]*[.\s]*最終輸出與自動化交接[^\n]*\n.*',
        '',
        content,
        flags=re.DOTALL
    )

    # 擷取 Mermaid 並準備下載
    pattern = re.compile(r'```mermaid\n(.*?)\n```', re.DOTALL)
    matches = pattern.findall(content)
    new_content = content

    print(f"[{base_name}] 開始自動建置 Word (DOCX) 企劃實體檔案...")
    if matches:
        print(f" => 發現了 {len(matches)} 個 Mermaid 圖表區塊，正在透過 API 生成手繪風圖片...", flush=True)
        for i, mermaid_text in enumerate(matches):
            url = create_api_url(mermaid_text.strip())
            output_png = f"{base_name}_chart_{i+1}.png"
            output_png_full = os.path.join(work_dir, output_png)
            
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=15) as response:
                    with open(output_png_full, 'wb') as img_file:
                        img_file.write(response.read())
                print(f"   -> 成功下載圖表：{output_png}", flush=True)
                
                # 將強制換頁的 openxml 語法塞在流程圖區塊上方
                page_break = '```{=openxml}\n<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```\n'
                original_block = f"```mermaid\n{mermaid_text}\n```"
                img_markdown = f"{page_break}![]({output_png})"
                new_content = new_content.replace(original_block, img_markdown)
                
            except Exception as e:
                print(f"   -> 轉換圖表失敗 (跳過圖表 {i+1}): {e}", flush=True)

    # 輸出過渡時期的一份包含圖片引用的 MD
    temp_md_name = f"{base_name}_temp_build.md"
    temp_md_full = os.path.join(work_dir, temp_md_name)
    with open(temp_md_full, 'w', encoding='utf-8') as f:
        f.write(new_content)

    # 呼叫 Pandoc 轉出 Word
    output_docx_name = f"{base_name}.docx"
    print(f" => 啟動 Pandoc 引擎：開始打包為 Word 規格書 ({output_docx_name})...", flush=True)
    
    try:
        # 尋找是否有客製化的 Word 範本 (用來控制縮排、段落不分頁等樣式)
        template_path = os.path.join(work_dir, "word_template.docx")
        
        # 指定在 md 所在資料夾執行，確保相對路徑抓得到圖片
        cmd = ["pandoc", temp_md_name, "-o", output_docx_name]
        
        # 如果使用者有自訂範本，就套用進去
        if os.path.exists(template_path):
            cmd.extend(["--reference-doc", template_path])
            print(f"   -> 偵測到自訂排版範本 (word_template.docx)，自動套用樣式...", flush=True)
        result = subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            print(f" => 大功告成！Word 檔已完美儲存於：\n   {os.path.join(work_dir, output_docx_name)}", flush=True)
        else:
            print(f" => Pandoc 發生例外狀況：{result.stderr}", flush=True)
    except FileNotFoundError:
        print(" => [警告] 未偵測到 Pandoc 工具！要產生 Word，電腦必須安裝 Pandoc (https://pandoc.org/installing.html)。", flush=True)
    except Exception as e:
        print(f" => 發生不明錯誤: {e}", flush=True)

    # 任務結束，不論成功或失敗都把帶圖片引用的過渡期 MD 殺掉，保持資料夾整潔
    if os.path.exists(temp_md_full):
        os.remove(temp_md_full)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("使用教學: python build_docx.py <你要打包的規格書檔案.md>")
    else:
        build(sys.argv[1])
