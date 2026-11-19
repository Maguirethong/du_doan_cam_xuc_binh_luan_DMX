import re
import html
import unicodedata
from collections import Counter
import pandas as pd
from pyvi import ViTokenizer

for i in range(1,6):

    input_filepath = r"Comment_cleaned\comments_{}.xlsx".format(i)      
    output_filepath = r"Comment_cleaned\comments_{}_loc_cau.xlsx".format(i)
    xls = pd.ExcelFile(input_filepath)
    df_old = pd.read_excel(xls, sheet_name="Comments")
    df = df_old[["STT", "Comment","Nhãn"]]

    def clean_basic_for_words(text: str) -> str:
        if pd.isna(text) or not isinstance(text, str):
            return ""
        text = text.lower()
        text = unicodedata.normalize("NFC", text).lower()
        text = re.sub(r"<[^>]+>", " ", text)
        text = re.sub(r"https?://\S+|www\.\S+|ftp://\S+", " ", text)
        text = re.sub(r"(.)\1{2,}", r"\1", text)  # tốtttt -> tốt ; \1 tham chieu phan tu group 1 (co lap lai phan tu do khong)
        text = re.sub(r"[^\w\s]", " ", text)
        text = re.sub(r"\b(ĐMX|đmx)\b", "dmx", text)
        return ViTokenizer.tokenize(text) if text else ""


    stop_word_temp = {
        "rất", "rat", "quá", "qua", "cực", "cực_kỳ", "siêu", "lắm", "lam", "khá", "thật", "thật_sự", "cũng",
        "dùng", "xài", "sài", "sử_dụng", "sữ_dụng", "sd", "dung", "xai", "sai", "mua", "thấy", "cảm_nhận",
        "sản_phẩm", "sp", "hàng", 
        "nha", "nhé", "nhe", "ạ", "a", "à", "ơi", "luôn", "lun", "rồi", "nữa", "thôi",
        "thì", "là", "và", "của", "có", "cho", "với", "này", "đó", "như", "ở", "về", "đến", "nay"
    }

    word_reference = {
        "tot": "tốt", "tôt": "tốt", "good": "tốt",
        "oke": "ok", "oki": "ok", "okii": "ok", "okie": "ok", "okay": "ok", "okela": "ok",
        "on": "ổn", "ổn_áp": "ổn",
        "bt": "bình_thường", "bthuong": "bình_thường",
        "ko": "không", "k": "không", "kg": "không", "hk": "không", "hok": "không", "chả": "không", "chẳng": "không",
        "dc": "được", "đc": "được", "duoc": "được",
        "ĐMX" : "dmx", "đmx" :"dmx"
    }

    def create_temp_comparison_key(text):
        cleaned = clean_basic_for_words(text)
        if not cleaned:
            return ""
        # chuan hoa tu viet tat/ dong nghia
        tokens = [word_reference.get(w, w) for w in cleaned.split()]
        reduced_tokens = [w for w in tokens if w not in stop_word_temp] # loai bo tu

        return " ".join(reduced_tokens)
    
    def tokenize_sentence(text):
        sentences = re.split(r'(?<=[.!?])\s*|(?<=[a-zà-ỹ0-9])(?=[A-Z])', text)

        sentences = [s.strip() for s in sentences if s.strip()]

        return " ".join(sentences)

    df["Temp_Key"] = df["Comment"].apply(create_temp_comparison_key)
    df_valid = df[df["Temp_Key"] != ""].copy()
    df_valid["Comment"] = df_valid["Comment"].apply(tokenize_sentence)

    df_unique_original = df_valid.drop_duplicates(subset=["Temp_Key"], keep="first").copy()

    print(f"\nSố câu hợp lệ trước khi lọc trùng : {len(df_valid)}")
    print(f"Số câu gốc duy nhất sau khi lọc   : {len(df_unique_original)}")
    print(f"Số câu giống nhau đã bị loại bỏ   : {len(df_valid) - len(df_unique_original)}")

    df_final = df_unique_original[["STT", "Comment", "Nhãn"]].reset_index(drop=True)
    df_final["STT"] = df_final.index + 1  # Đánh lại STT liên tục từ 1

    with pd.ExcelWriter(output_filepath, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="Comments", index=False)

    print(f"\nĐã lưu {len(df_final)} câu gốc duy nhất vào file '{output_filepath}' (sheet 'Comments').")