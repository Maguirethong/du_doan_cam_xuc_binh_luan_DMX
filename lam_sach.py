import re
import unicodedata
import pandas as pd
from pyvi import ViTokenizer

def clean_text_pipeline(text):
    if pd.isna(text) or not isinstance(text, str):
        return ""
    
    text = unicodedata.normalize("NFC", text)# Chuẩn hóa Unicode
    text = re.sub(r"<[^>]+>", " ", text) #loai html
    text = re.sub(r"https?://\S+|www\.\S+|ftp://\S+", " ", text)#loai link
    text = re.sub(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', '', text)#loai ngay
    text = re.sub(r"[^\w\s]", " ", text) #loai bo ki tu dac biet
    text = text.replace("_", " ")
    text = re.sub(r'\d+', '', text)
    text = re.sub(r"\b(ko|k|kg|khg|hk|hok|hong)\b", "không", text)
    text = re.sub(r"\b(dc|đc|dk|đk)\b", "được", text)
    text = re.sub(r"\b(sp)\b", "sản phẩm", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = text.lower()
    if text:
        text = re.sub(r"\bkhông biết\b", "không_biết", text)
        tokenized_text = ViTokenizer.tokenize(text) #tach tu
        return remove_stopwords(tokenized_text)
    
def remove_stopwords(tokenized_text):
    stopwords = {
        "là", "và", "của", "có", "những", "các", "được", "trong", "cho", "với",
        "này", "đó", "khi", "mà", "thì", "để", "từ", "về", "như", "ra", "vào",
        "lên", "xuống", "theo", "tại", "do", "bởi", "vì", "nếu", "sẽ", "đang",
        "đã", "vẫn", "cũng", "nhé", "nha", "ạ", "à", "ơi", "ừ", "ha", "hihi",
        "haha", "luôn", "rồi", "thôi", "nữa", "cái", "chiếc", "con", "bộ", "loại",
        "mình", "tôi", "em", "anh", "chị", "bạn", "shop", "cửa_hàng", "nhà", "người",
        "lần", "ngày", "tháng", "năm", "giờ", "hôm", "nay", "qua", "trước", "sau",
        "trên", "dưới", "ngoài", "trong", "đây", "đấy", "ấy", "nào", "gì", "sao",
        "ai", "đâu", "vậy", "thế", "nên", "phải", "cần", "muốn", "biết", "thấy"
    }

    giu_lai = {
        "không", "chưa", "chẳng", "chả", "đừng", #phu dinh
        "rất", "quá", "cực", "cực_kỳ", "khá", "hơi", "tạm", "bình_thường", #muc do
        "nhưng", "tuy_nhiên", "bị", "mất", "thiếu", "lỗi", "hư", "hỏng" #tieu cuc
    }
    final_stopwords = stopwords - giu_lai
    tokens = tokenized_text.split()
    filtered_tokens = [word for word in tokens if word not in final_stopwords]

    return " ".join(filtered_tokens)

if __name__ == "__main__":
    for i in range(1,6):
        input_filepath = r"Comment_cleaned\comments_{}_loc_cau.xlsx".format(i)
        output_filepath = r"Comment_cleaned\comments_{}_cleaned.xlsx".format(i)
        xlsx = pd.ExcelFile(input_filepath)
        df= pd.read_excel(xlsx, sheet_name="Comments")
        df_copy = df.copy()
        df_copy["Comment"] = df["Comment"].apply(clean_text_pipeline)
        df_final = df_copy[df_copy["Comment"] != ""].copy()
        with pd.ExcelWriter(output_filepath, engine="openpyxl") as ew:
            df_copy.to_excel(ew, sheet_name="Comments", index=False)