import re
import os


def extract_exercises(input_file, output_file):
    if not os.path.exists(input_file):
        print(f"错误：找不到源文件 {input_file}")
        return

    with open(input_file, "r", encoding="utf-8") as f:
        content = f.read()

    # 改进后的匹配模式：
    # \{习题演练\} : 匹配起始标记
    # .*? : 非贪婪匹配习题内容
    # (?= ... ) : 正向肯定断言，匹配到以下内容之一时停止：
    #   1. \n\s*第[一二三四五六七八九十]+[部分章节] : 匹配“第一部分”、“第二章”、“第三节”等
    #   2. \n\s*[一二三四五六七八九十]+[、.] : 匹配“一、”、“二、”等一级标题
    #   3. \n\s*[\(（][一二三四五六七八九十]+[\)）] : 匹配“（一）”、“（二）”等二级标题
    #   4. |\{习题演练\} : 下一个习题演练块
    #   5. |$ : 文件末尾
    stop_pattern = (
        r"\n\s*(?:第[一二三四五六七八九十]+[部分章节])|"
        r"\n\s*[一二三四五六七八九十]+[、.]|"
        r"\n\s*[\(（][一二三四五六七八九十]+[\)）]|"
        r"\{习题演练\}|$"
    )

    pattern = r"\{习题演练\}.*?(?=" + stop_pattern + ")"

    matches = re.findall(pattern, content, re.DOTALL)

    with open(output_file, "w", encoding="utf-8") as f:
        for i, match in enumerate(matches, 1):
            # 清理片段内容，去除多余的空白
            cleaned_match = match.strip()
            if cleaned_match:
                f.write(f"--- 提取片段 {i} ---\n")
                f.write(cleaned_match)
                f.write("\n\n")

    print(f"提取完成，共找到 {len(matches)} 处习题演练。结果已保存至 {output_file}")


if __name__ == "__main__":
    extract_exercises("2025柯凡社区讲义.txt", "习题演练汇总.txt")
