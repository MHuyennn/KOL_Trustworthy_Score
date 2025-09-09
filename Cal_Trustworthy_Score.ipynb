
import pandas as pd
import numpy as np

posts_df = pd.read_excel('Raw_Data.xlsx', sheet_name='Posts')
profiles_df = pd.read_excel('Raw_Data.xlsx', sheet_name='Profiles')
scandals_df = pd.read_excel('Raw_Data.xlsx', sheet_name='Scandals')
comments_df = pd.read_excel('Raw_Data.xlsx', sheet_name='Comments')

# Chuẩn hóa tên cột
posts_df.columns = posts_df.columns.str.lower()
profiles_df.columns = profiles_df.columns.str.lower()
scandals_df.columns = scandals_df.columns.str.lower()
comments_df.columns = comments_df.columns.str.lower()

# Bước 1: Kiểm tra KOL và bài đăng
print("Kiểm tra KOL và bài đăng")
kol_posts = posts_df.groupby('kol_id')['post_id'].apply(list).reset_index()
print("Danh sách KOL và Post_ID:")
print(kol_posts)

# Kiểm tra cụ thể cho K02, K12, K13
target_kols = ['K02', 'K12', 'K13']
for kol in target_kols:
    kol_data = posts_df[posts_df['kol_id'] == kol]
    if len(kol_data) > 0:
        print(f"\nBài đăng của {kol}:")
        print(kol_data[['post_id', 'professionalism_score', 'is_expertise_match']])
    else:
        print(f"\nKhông tìm thấy bài đăng nào cho {kol}")

# Bước 2: Kiểm tra bình luận cho mỗi bài đăng
print("\nKiểm tra bình luận")
post_comments_count = comments_df.groupby('post_id').size().reset_index(name='comment_count')
print("Số lượng bình luận cho mỗi Post_ID:")
print(post_comments_count)

# Bước 3: Tính Authenticity_Score và Community_Trust_Score
def calculate_authenticity_and_community_scores(post_id, comments_df, is_expertise_match):
    post_comments = comments_df[comments_df['post_id'] == post_id]
    total_comments = len(post_comments)

    if total_comments == 0:
        print(f"Post_ID {post_id}: Không có bình luận")
        return 0, 0

    # Tính On_Topic_Ratio
    on_topic_types = ['user_experience_share', 'conversion_intent', 'inquiry_product', 'brand_loyalty']
    on_topic_comments = post_comments[post_comments['comment_type'].isin(on_topic_types)]
    on_topic_ratio = len(on_topic_comments) / total_comments if total_comments > 0 else 0

    # Tính Other_Product_Ratio
    other_product_comments = post_comments[post_comments['comment_type'] == 'other_product_mention']
    other_product_ratio = len(other_product_comments) / total_comments if total_comments > 0 else 0

    # Tính Focus_Factor và Authenticity_Score
    focus_factor = on_topic_ratio - other_product_ratio
    authenticity_score = (5 if is_expertise_match else 1) + (max(0, focus_factor) * 5)

    # Tính Community_Trust_Score
    positive_comments = post_comments[post_comments['sentiment'] == 1]
    negative_comments = post_comments[post_comments['sentiment'] == -1]

    weighted_positive_trust = sum(positive_comments['is_high_quality'] * positive_comments['sentiment'] * positive_comments['like'])
    weighted_negative_trust = sum(negative_comments['is_high_quality'] * negative_comments['sentiment'] * negative_comments['like'])
    total_relevant_comments = len(positive_comments) + len(negative_comments)

    if total_relevant_comments > 0:
        base_score = ((weighted_positive_trust - weighted_negative_trust) / total_relevant_comments) * 10
        community_trust_score = max(0, min(10, base_score))
    else:
        print(f"Post_ID {post_id}: Không có bình luận tích cực/tiêu cực")
        community_trust_score = 0

    print(f"Post_ID {post_id}: Authenticity_Score={authenticity_score:.2f}, Community_Trust_Score={community_trust_score:.2f}")
    return authenticity_score, community_trust_score

# Áp dụng tính toán cho mỗi bài đăng
posts_df[['authenticity_score', 'community_trust_score']] = posts_df.apply(
    lambda row: calculate_authenticity_and_community_scores(row['post_id'], comments_df, row['is_expertise_match']),
    axis=1, result_type='expand'
)

# Bước 4: Tính Post_Trust_Score
posts_df['post_trust_score'] = (0.3 * posts_df['professionalism_score'] +
                                0.3 * posts_df['authenticity_score'] +
                                0.4 * posts_df['community_trust_score'])

# In chi tiết điểm số cho mỗi bài đăng
print("\nĐiểm số chi tiết cho mỗi bài đăng")
print(posts_df[['post_id', 'kol_id', 'professionalism_score', 'authenticity_score', 'community_trust_score', 'post_trust_score']])

# Bước 5: Tính Average_Post_Trust_Score
avg_post_trust = posts_df.groupby('kol_id')['post_trust_score'].mean().reset_index()
avg_post_trust.columns = ['kol_id', 'average_post_trust_score']
print("\nAverage_Post_Trust_Score cho mỗi KOL")
print(avg_post_trust)

# Bước 6: Tính Total_Risk_Penalty
scandals_df['resolution_multiplier'] = scandals_df['resolved'].apply(lambda x: 0.4 if x == 1 else 1.0)
scandals_df['scandal_penalty_point'] = scandals_df['severity_score'] * scandals_df['resolution_multiplier']
total_risk_penalty = scandals_df.groupby('kol_id')['scandal_penalty_point'].sum().reset_index()
total_risk_penalty.columns = ['kol_id', 'total_risk_penalty']
print("\nTotal_Risk_Penalty cho mỗi KOL")
print(total_risk_penalty)

# Bước 7: Kết hợp tất cả KOL từ Profiles
kol_scores = profiles_df[['kol_id', 'kol_name']].merge(avg_post_trust, on='kol_id', how='left')
kol_scores = kol_scores.merge(total_risk_penalty, on='kol_id', how='left')

# Xử lý NaN
kol_scores['average_post_trust_score'] = kol_scores['average_post_trust_score'].fillna(0)
kol_scores['total_risk_penalty'] = kol_scores['total_risk_penalty'].fillna(0)
kol_scores['total_risk_penalty'] = kol_scores['total_risk_penalty']*0.2

# Bước 8: Tính Trustworthy_Score
kol_scores['trustworthy_score'] = kol_scores.apply(
    lambda row: max(0, row['average_post_trust_score'] - row['total_risk_penalty']), axis=1)

# Bước 9: Sắp xếp theo Trustworthy_Score giảm dần
kol_scores = kol_scores.sort_values(by='trustworthy_score', ascending=False)

# Bước 10: In kết quả cuối cùng
print("\nKết quả Trustworthy_Score")
print(kol_scores[['kol_id', 'kol_name', 'trustworthy_score', 'average_post_trust_score', 'total_risk_penalty']])

# Lưu kết quả ra file Excel
kol_scores[['kol_id', 'kol_name', 'trustworthy_score']].to_excel('KOL_Trustworthy_Scores.xlsx', index=False)
