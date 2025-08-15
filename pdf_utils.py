from PIL import Image
import os
from tkinter import messagebox
import fitz  # PyMuPDF
import easyocr

def create_pdf_from_images(image_dir: str, pdf_path: str):
    images = []
    image_files = sorted([f for f in os.listdir(image_dir) if f.endswith((".png", ".jpg", ".jpeg"))])

    if not image_files:
        messagebox.showinfo("정보", "PDF로 변환할 이미지가 없습니다.")
        return

    for file_name in image_files:
        image_path = os.path.join(image_dir, file_name)
        img = Image.open(image_path)
        if img.mode == 'RGBA':
            img = img.convert('RGB')
        images.append(img)

    if images:
        images[0].save(pdf_path, save_all=True, append_images=images[1:])
        messagebox.showinfo("성공", f"PDF 파일이 저장되었습니다: {pdf_path}")


def create_searchable_pdf(image_dir: str, pdf_path: str):
    try:
        messagebox.showinfo("정보", "OCR 작업은 시간이 걸릴 수 있습니다. 잠시만 기다려주세요...")
        reader = easyocr.Reader(['ko', 'en']) 
        
        image_files = sorted([f for f in os.listdir(image_dir) if f.endswith((".png", ".jpg", ".jpeg"))])
        if not image_files:
            messagebox.showinfo("정보", "PDF로 변환할 이미지가 없습니다.")
            return

        doc = fitz.open()

        for file_name in image_files:
            image_path = os.path.join(image_dir, file_name)
            
            # OCR 실행
            result = reader.readtext(image_path)
            
            # PyMuPDF로 이미지 열기 및 페이지 생성
            img_doc = fitz.open(image_path)
            page = doc.new_page(width=img_doc[0].rect.width, height=img_doc[0].rect.height)
            page.insert_image(page.rect, filename=image_path)

            # 보이지 않는 텍스트 추가
            for (bbox, text, prob) in result:
                # bbox는 [[x1, y1], [x2, y1], [x2, y2], [x1, y2]] 형식
                x1, y1 = bbox[0]
                x2, y2 = bbox[2]
                rect = fitz.Rect(x1, y1, x2, y2)
                page.insert_text(rect.bottom_left, text, render_mode=3, fontsize=8)

        doc.save(pdf_path, garbage=4, deflate=True, clean=True)
        doc.close()
        messagebox.showinfo("성공", f"검색 가능한 PDF 파일이 저장되었습니다: {pdf_path}")

    except Exception as e:
        messagebox.showerror("OCR 오류", f"OCR 처리 중 오류가 발생했습니다: {e}")
