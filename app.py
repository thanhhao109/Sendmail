import os
import pandas as pd
import pythoncom
import win32com.client as win32
from flask import Flask, render_template, request

app = Flask(__name__)
from flask import send_from_directory

@app.route('/uploads/<path:filename>')
def uploaded_file(filename):
    return send_from_directory(UPLOAD_DIR, filename)


# ===============================
#   THƯ MỤC LƯU FILE
# ===============================
UPLOAD_DIR = "uploads"
EXCEL_DIR = "excel_uploads"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)


# ===============================
#   HÀM GỬI EMAIL: PNG INLINE + PDF
# ===============================
def send_email_png_inline_and_pdf(to_email, subject, png_path, pdf_path=None):
    try:
        pythoncom.CoInitialize()

        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.SentOnBehalfOfName = "customerservice@ocbs.com.vn"

        mail.To = to_email
        mail.Subject = subject

        # CID để hiển thị PNG inline
        cid = "img001"

        # HTML email chỉ chứa PNG lớn
        mail.HTMLBody = f"""
        <html>
        <body style="margin:0; padding:0;">
            <div style="text-align:center;">
                <img src="cid:{cid}" style="width:100%; max-width:100%; height:auto; display:block;">
            </div>
        </body>
        </html>
        """

        # Gắn PNG dạng inline
        att_png = mail.Attachments.Add(os.path.abspath(png_path))
        att_png.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
            cid
        )

        # Gắn PDF nếu có
        if pdf_path:
            mail.Attachments.Add(os.path.abspath(pdf_path))

        mail.Send()
        print("✔ Đã gửi:", to_email)

    except Exception as e:
        print("❌ Lỗi gửi email:", e)

    finally:
        pythoncom.CoUninitialize()



# ===============================
#   GỬI TỪNG EMAIL
# ===============================
@app.route("/", methods=["GET", "POST"])
def index():
    png_files = [f for f in os.listdir(UPLOAD_DIR) if f.lower().endswith(".png")]
    pdf_files = [f for f in os.listdir(UPLOAD_DIR) if f.lower().endswith(".pdf")]

    if request.method == "POST":
        subject = request.form["subject"]     # <-- Lấy Subject từ form
        email = request.form["email"]
        png_name = request.form["filename"]
        pdf_name = request.form.get("pdf_file")

        png_path = os.path.join(UPLOAD_DIR, png_name)
        pdf_path = os.path.join(UPLOAD_DIR, pdf_name) if pdf_name else None

        send_email_png_inline_and_pdf(
            to_email=email,
            subject=subject,
            png_path=png_path,
            pdf_path=pdf_path
        )

        return "✔ Đã gửi email thành công!"

    return render_template("form.html", png_files=png_files, pdf_files=pdf_files)



# ===============================
#   GỬI EMAIL HÀNG LOẠT
# ===============================
@app.route("/bulk", methods=["GET", "POST"])
def bulk():
    if request.method == "POST":

        # Lấy tiêu đề email từ form
        subject = request.form.get("subject", "OCBS - Thông báo Kết quả mua IPO")

        # Nhận file Excel
        excel_file = request.files["excel"]
        excel_path = os.path.join(EXCEL_DIR, excel_file.filename)
        excel_file.save(excel_path)

        # Đọc Excel
        try:
            df = pd.read_excel(excel_path)
        except Exception:
            return "❌ Excel lỗi hoặc thiếu openpyxl. Chạy lệnh: pip install openpyxl"

        # Kiểm tra đúng cấu trúc cột
        if not all(col in df.columns for col in ["Email", "PNG", "PDF"]):
            return "❌ Excel phải có 3 cột: Email – PNG – PDF"

        sent_count = 0

        for _, row in df.iterrows():
            email = str(row["Email"]).strip()
            png_name = str(row["PNG"]).strip()
            pdf_name = str(row["PDF"]).strip()

            png_path = os.path.join(UPLOAD_DIR, png_name)
            pdf_path = os.path.join(UPLOAD_DIR, pdf_name)

            if not os.path.isfile(png_path):
                print("❌ Không tìm thấy PNG:", png_path)
                continue

            if not os.path.isfile(pdf_path):
                print("❌ Không tìm thấy PDF:", pdf_path)
                continue

            send_email_png_inline_and_pdf(
                to_email=email,
                subject=subject,   # <-- Áp dụng subject nhập ở bulk
                png_path=png_path,
                pdf_path=pdf_path
            )

            sent_count += 1

        return f"✔ Đã gửi thành công {sent_count} email!"

    return render_template("bulk.html")



# ===============================
#   CHẠY APP – KHÔNG DEBUG
# ===============================
if __name__ == "__main__":
    app.run(debug=False, port=5005)
