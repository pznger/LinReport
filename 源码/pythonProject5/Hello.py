import os
# sk-JGR1yG7xV86kqxDcty03hwEAWhXhaVnwcxEKGU0Fp40rAbgs
def combine_java_files_to_txt(base_directory, output_txt_path):
    # 确保输出文件的路径存在
    os.makedirs(os.path.dirname(output_txt_path), exist_ok=True)

    # 打开输出文件
    with open(output_txt_path, 'w', encoding='utf-8') as output_file:
        # 遍历给定的目录及其所有子目录
        for root, dirs, files in os.walk(base_directory):
            for filename in files:
                if filename.endswith('.java'):
                    file_path = os.path.join(root, filename)
                    # 打开并读取每个.java文件的内容
                    with open(file_path, 'r', encoding='utf-8') as java_file:
                        # 将文件内容写入输出文件
                        output_file.write(f'--- {filename} from {root} ---\n')
                        output_file.write(java_file.read())
                        output_file.write('\n\n')

# 使用示例
base_directory = 'G:\\HeadFirst-main\\HeadFirst-java\\command'  # 替换为你的.java文件所在的顶级目录路径
output_txt_path = 'D:\\combined_java_files.txt'  # 替换为你想要保存.txt文件的路径

combine_java_files_to_txt(base_directory, output_txt_path)