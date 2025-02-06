import csv

def save_to_csv(file_path, data):
    """입력 데이터를 CSV 파일로 저장"""
    try:
        with open(file_path, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(data)
        print(f"데이터가 성공적으로 저장되었습니다: {file_path}")
    except Exception as e:
        print(f"파일 저장 중 오류 발생: {e}")

# 사용 예시
user_data = ["홍길동", "2025-02-01", "2025-02-10", "2025-02-15"]
save_to_csv("user_data.txt", user_data)