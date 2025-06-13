import os

folder_path = "GaN"  # 当前文件夹路径（可替换为绝对路径，如 "D:/data"）

# 获取文件夹下所有文件和文件夹名称
all_entries = os.listdir(folder_path)

# 过滤出文件（排除文件夹）
for entry in all_entries:
    if os.path.isfile(os.path.join(folder_path, entry)):
        print(entry)

