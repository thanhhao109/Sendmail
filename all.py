import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

from flask import Flask, render_template, request

app = Flask(__name__)

# ================================
# SMTP OFFICE365 CONFIG
# ================================
SMTP_EMAIL = "haont@ocbs.com.vn"      # ← sửa tại đây
SMTP_PASSWORD = "Thanhhao@109"        # ← sửa tại đây
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

# ================================
# FOLDER CONFIG
# ================================
UPLOAD_DIR = "uploads"
EXCEL_DIR = "excel_uploads"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)


# ==========================================================
# HÀM GỬI EMAIL BẰNG SMTP OFFICE365
# ==========================================================
def send_email_smtp(to_email, subject, png_path, pdf_path=None):
    print("Đang gửi đến:", to_email)

    msg = MIMEMultipart('related')
    msg['From'] = SMTP_EMAIL
    msg['To'] = to_email
    msg['Subject'] = subject

    # HTML body có PNG inline
    cid = "img001"
    html_body = f"""
    <html>
    <body style="margin:0;padding:0;">
        <div style="text-align:center;">
            <img src="cid:{cid}" style="width:100%;max-width:100%;height:auto;">
        </div>
    </body>
    </html>
    """
    msg.attach(MIMEText(html_body, 'html'))

    # Attach PNG inline
    with open(png_path, 'rb') as f:
        img = MIMEImage(f.read())
        img.add_header('Content-ID', f'<{cid}>')
        img.add_header("Content-Disposition", "inline")
        msg.attach(img)

    # Attach PDF (optional)
    if pdf_path:
        with open(pdf_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={os.path.basename(pdf_path)}"
            )
            msg.attach(part)

    # Gửi SMTP
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(SMTP_EMAIL, SMTP_PASSWORD)
    server.send_message(msg)
    server.quit()

    print("✔ Gửi thành công →", to_email)



# ==========================================================
# GỬI 1 EMAIL
# ==========================================================
@app.route("/", methods=["GET", "POST"])
def index():
    png_files = [f for f in os.listdir(UPLOAD_DIR) if f.lower().endswith(".png")]
    pdf_files = [f for f in os.listdir(UPLOAD_DIR) if f.lower().endswith(".pdf")]

    if request.method == "POST":
        subject = request.form["subject"]
        email = request.form["email"]
        png_name = request.form["filename"]
        pdf_name = request.form.get("pdf_file")

        png_path = os.path.join(UPLOAD_DIR, png_name)
        pdf_path = os.path.join(UPLOAD_DIR, pdf_name) if pdf_name else None

        send_email_smtp(
            to_email=email,
            subject=subject,
            png_path=png_path,
            pdf_path=pdf_path
        )

        return "✔ Đã gửi email thành công!"

    return render_template("form.html", png_files=png_files, pdf_files=pdf_files)



# ==========================================================
# GỬI HÀNG LOẠT TỪ EXCEL
# ==========================================================
@app.route("/bulk", methods=["GET", "POST"])
def bulk():
    if request.method == "POST":

        excel_file = request.files["excel"]
        excel_path = os.path.join(EXCEL_DIR, excel_file.filename)
        excel_file.save(excel_path)

        try:
            df = pd.read_excel(excel_path)
        except Exception:
            return "❌ File Excel lỗi hoặc thiếu openpyxl"

        # Excel cần có ít nhất 3 cột
        required_cols = ["Email", "PNG", "PDF"]
        for col in required_cols:
            if col not in df.columns:
                return f"❌ Thiếu cột: {col}"

        # Nếu Excel có cột Subject → dùng riêng từng dòng
        has_subject = "Subject" in df.columns

        sent_count = 0

        for _, row in df.iterrows():
            email = str(row["Email"]).strip()
            png_name = str(row["PNG"]).strip()
            pdf_name = str(row["PDF"]).strip()

            png_path = os.path.join(UPLOAD_DIR, png_name)
            pdf_path = os.path.join(UPLOAD_DIR, pdf_name) if pdf_name else None

            subject = row["Subject"] if has_subject else "OCBS – Thông báo kết quả giao dịch"

            if not os.path.isfile(png_path):
                print("❌ Không tìm thấy PNG:", png_path)
                continue

            if pdf_name and not os.path.isfile(pdf_path):
                print("❌ Không tìm thấy PDF:", pdf_path)
                continue

            send_email_smtp(
                to_email=email,
                subject=subject,
                png_path=png_path,
                pdf_path=pdf_path
            )

            sent_count += 1

        return f"✔ Đã gửi {sent_count} email thành công!"

    return render_template("bulk.html")



# ==========================================================
# CHẠY APP
# ==========================================================
if __name__ == "__main__":
    app.run(debug=True, port=5003)
