#coding=utf-8
from wordcloud import WordCloud
import jieba
from PIL import Image
import matplotlib.pyplot as plt
import numpy as np
import xlwt

def wordcloudplot(txt):
    path = 'e:/python/msyh.ttf'
    path = unicode(path, 'utf8').encode('gb18030')
    img = Image.new("RGB",(1024,1024))
    alice_mask = np.array(img)
    wordcloud = WordCloud(font_path=path,
                          background_color="white",
                          margin=5, width=1800, height=800, mask=alice_mask, max_words=2000, max_font_size=60,
                          random_state=42)
    wordcloud = wordcloud.generate(txt)
    wordcloud.to_file('e:/python/output/she2.jpg')
    plt.imshow(wordcloud)
    plt.axis("off")
    plt.show()

def stat(words):
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet("wordCount")  # Excel单元格名字
    key_list = []
    word_dict={}
    for word in words:
        if word not in word_dict:
            word_dict[word] = 1
        else:
            word_dict[word] += 1
    orderList = list(word_dict.values())
    # orderList.sort(reverse=True)
    key_list = list(word_dict.keys())
    for i in range(len(key_list)):
        sheet.write(i, 1, label=orderList[i])
        sheet.write(i, 0, label=key_list[i])
    wbk.save('output/wordCount.xls')  # 保存为 wordCount.xls文件


def main():
    a = []
    f = open(r'e:/python/resources/test-0.txt', 'r').read()
    words = list(jieba.cut(f))
    stat_words = []
    for word in words:
        if len(word) > 1:
            stat_words.append(word)
            a.append(word)
    txt = r' '.join(a)
    stat(stat_words)
    wordcloudplot(txt)

if __name__ == '__main__':
    main()