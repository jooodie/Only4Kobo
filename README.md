# PDF 轉 EPUB / KEPUB 工具

將 PDF 電子書轉換為 EPUB，並自動再轉為 Kobo 較穩定的 KEPUB 格式。

本工具使用 [PyMuPDF](https://pymupdf.readthedocs.io/) 解析 PDF 文字區塊，自動清理頁首、頁尾與頁碼，依字體大小辨識標題層級，再以 [EbookLib](https://github.com/aerkalov/ebooklib) 封裝成 EPUB，最後透過 Calibre CLI 轉成 KEPUB。

## 功能特色

- 自動提取 PDF 第一頁作為 EPUB 封面
- 過濾頁首、頁尾與常見頁碼格式
- 依字體大小自動辨識 h1～h3 標題
- 合併相鄰段落，還原閱讀流暢度
- 依 h1 標題或固定區塊數自動分章
- 支援自訂書名、作者、語言等中繼資料
- 自動呼叫 `ebook-convert` 產生 `.kepub.epub`
- 自動從 metadata/檔名抓作者並歸檔（大小寫統一，找不到則 `Unknown`）

## 環境需求

- Python 3.10 以上
- Calibre（需可在終端執行 `ebook-convert`）
- Python 套件（見 `requirements.txt`）

## 安裝

```bash
git clone https://github.com/jooodie/pdf2epub.git
cd pdf2epub

python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

若 `ebook-convert` 指令不存在，請先安裝 Calibre，並確認下列指令可執行：

```bash
ebook-convert --version
```

## 使用流程（建議照此順序）

1. 進入專案資料夾
2. 啟用虛擬環境
3. 執行轉換指令
4. 取得 `*.epub` 與 `*.kepub.epub`
5. 到作者資料夾取檔（找不到作者時在 `Unknown/`）

```bash
cd /Users/jodie/Desktop/Jooo/For_Kobo
source .venv/bin/activate
python pdf_to_epub.py "輸入檔.pdf" "輸出檔.epub"
```

## 基本用法

```bash
python pdf_to_epub.py <輸入PDF> <輸出EPUB>
```

### 範例

```bash
python pdf_to_epub.py 快思慢想.pdf 快思慢想.epub
```

轉換完成後，建議將 `.kepub.epub` 傳到 Kobo 閱讀器。

## 命令列參數

| 參數 | 說明 | 預設值 |
|------|------|--------|
| `input_pdf` | 輸入 PDF 檔案路徑 | （必填） |
| `output_epub` | 輸出 EPUB 檔案路徑 | （必填） |
| `--title` | 電子書標題 | PDF 檔名 |
| `--author` | 作者名稱 | `Unknown` |
| `--language` | 語言代碼（如 `zh`、`en`） | `zh` |
| `--header-ratio` | 頁首區域高度比例（0～1） | `0.08` |
| `--footer-ratio` | 頁尾區域高度比例（0～1） | `0.08` |
| `--skip-kepub` | 只輸出 EPUB，不轉 KEPUB | 關閉 |

### 進階範例

```bash
python pdf_to_epub.py book.pdf book.epub \
  --title "快思慢想" \
  --author "丹尼爾·卡尼曼" \
  --language zh
```

若 PDF 的頁首或頁尾較高，可調整比例以改善過濾效果：

```bash
python pdf_to_epub.py book.pdf book.epub --header-ratio 0.12 --footer-ratio 0.10
```

## 轉換流程

預設會依序完成以下六個步驟：

```
[1/6] 提取封面：將 PDF 第 1 頁渲染為封面圖片
[2/6] 解析 PDF 內文：讀取文字區塊（跳過封面頁）
[3/6] 清理結構並辨識標題：過濾頁首/頁尾、分類標題、合併段落
[4/6] 封裝 EPUB：產生含封面、目錄、章節的 EPUB 檔案
[5/6] Calibre 轉換：以 ebook-convert 產生 `.kepub.epub`
[6/6] 作者歸檔：依作者建立資料夾並複製 `.kepub.epub`
```

若只想輸出 EPUB，可加上：

```bash
python pdf_to_epub.py book.pdf book.epub --skip-kepub
```

## 適用與限制

### 適用的 PDF

- 有**文字層**的 PDF（例如從 Word、LaTeX 匯出的電子書）
- 第一頁可作為封面（內文從第 2 頁開始解析，避免重複）

### 不適用的 PDF

- **掃描版 PDF**（整頁為圖片、無文字層）— 本工具無法辨識圖片中的文字
- 文字以圖片嵌入的 PDF
- 複雜排版（多欄、表格、公式）可能無法完美還原

若轉換時出現「未從 PDF 提取到可用文字」錯誤，通常表示 PDF 沒有可提取的文字層。

## 授權

本專案供個人使用。轉換受版權保護的書籍時，請遵守相關法律規定。
