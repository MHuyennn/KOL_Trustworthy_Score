import pandas as pd

# Đọc file Excel với tất cả các sheet
excel_file = "KOL.xlsx"
sheets = pd.read_excel(excel_file, sheet_name=["Profiles", "Products", "Posts", "Comments", "Scandals"])

# Gán các DataFrame từ các sheet
profiles_df = sheets["Profiles"]
products_df = sheets["Products"]
posts_df = sheets["Posts"]
comments_df = sheets["Comments"]
scandals_df = sheets["Scandals"]

# Hợp nhất các sheets
merged_df = posts_df.merge(profiles_df, on="KOL_ID", how="left")
merged_df = merged_df.merge(products_df, on="Product_ID", how="left")
merged_df = merged_df.merge(comments_df, on="Post_ID", how="left")
merged_df = merged_df.merge(scandals_df, on="KOL_ID", how="left")

# Chuẩn hóa tên cột
merged_df.columns = merged_df.columns.str.replace(' ', '_').str.lower()

# Chuyển đổi các cột dạng boolean 
merged_df['is_expertise_match'] = merged_df['is_expertise_match'].map({1: True, 0: False, None: None})
merged_df['resolved'] = merged_df['resolved'].map({1: True, 0: False, None: None})

# Xử lý giá trị thiếu
string_columns = ['post_content', 'comment_content', 'description', 'kol_name', 'niche',
                 'product_name', 'product_category', 'scandal_type']
for col in string_columns:
    merged_df[col] = merged_df[col].fillna('')

# Lưu vào file CSV
output_csv = 'KOL_dataset.csv'
merged_df.to_csv(output_csv, index=False, encoding='utf-8')

# In thông báo hoàn thành
print(f"Dữ liệu từ 5 sheet đã được hợp nhất và lưu vào file '{output_csv}'.")
