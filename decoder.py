import os
import docx
import json

from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

from utils import Paragraph_, element_, Stack, Task, task_2_json

stack = Stack()
num = -1


def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def read_table(table):
    return [[cell.text for cell in row.cells] for row in table.rows]


def getUpId_no(paraId, style):
    global stack
    global num  # 适用于给正文编号，因为正文不会被压入栈中
    if stack.is_empty():
        if style != 0:
            num = -1  # 由于有新的标题，所以正文计数器清零
            stack.push(element_(paraId, style, 0))
            return [-1, 0]
        else:
            num += 1
            return [-1, num]
    # 若为标题
    if style != 0:
        while stack.top().style > style:
            stack.pop()
        num = -1  # 由于更换了新的标题，所以将正文的计数器清零
        if stack.top().style == style:
            no = stack.pop().no
            no = no + 1
        else:
            no = 0
        upParaId = stack.top().paraId
        stack.push(element_(paraId, style, no))
        return [upParaId, no]
    # 若为正文
    upParaId = stack.top().paraId
    num += 1
    return [upParaId, num]


def read_word(word_path):
    global stack

    # @yws
    articleTittle = ''  # 全文标题
    stuName = ''        # 学生姓名
    teacherName = ''    # 教师姓名
    majorName = ''      # 专业名称
    gradTime  = ''      # 毕业时间
    # @yws

    level = ''  # 正文-一级标题-二级标题-三级标题-四级标题-五级标题-列表 => 1,2,3,4,5,6
    module = ''  # 普通文字-摘要 => 1,2 =>默认为空，即普通文字
    list_type = ''  # 顺序-非顺序 => 1,2 =>默认为空，即不是列表模块
    language = '1'  # 中文-英文 => 1,2 =>默认为’1‘，即中文

    paras = []
    catalog = None  # 目录放在此处

    stack = Stack()
    i = 0

    paraId = -1
    doc = docx.Document(word_path)
    for block in iter_block_items(doc):

        # if block == 第一个 -> 标题
        if isinstance(block, Paragraph):
            content = block.text
            ac_content = "".join(content.split())   # 去除了空格，\t,\n等字符的文本，用于判断
            # 模块-语言 控制器
            if len(ac_content) == 0:
                continue
            elif ac_content == '摘要':
                module = '2'
                language = '1'
            elif content == 'ABSTRACT':
                module = '2'
                language = '2'
            elif content.split(" ")[0] == '第一章':
                module = ''
                language = '1'
            # @yws
            elif ac_content.startswith('题目') and articleTittle=='':
                articleTittle = ac_content[2:]
                continue
            elif ac_content.startswith('专业名称') and majorName=='':
                majorName = ac_content[4:]
                continue
            elif ac_content.startswith('学生姓名') and stuName=='':
                stuName = ac_content[4:]
                continue
            elif ac_content.startswith('指导教师') and teacherName=='':
                teacherName = ac_content[4:]
                continue
            elif ac_content.startswith('毕业时间') and gradTime=='':
                gradTime = ac_content[4:]
                continue
            # @yws
            style_name = block.style.name
            paraId += 1
            if style_name.startswith('Heading'):
                level = style_name.split(" ")[1]
                list_type = ''
                [upParaId, no] = getUpId_no(paraId, int(level))
                style = level + module + list_type + language
            elif style_name.startswith('List'):
                level = '6'
                list_type = '1'
                [upParaId, no] = getUpId_no(paraId, 0)
                # style = style_name.split(" ")[1]
                style = level + module + list_type + language
            else:
                level = '1'
                list_type = ''
                [upParaId, no] = getUpId_no(paraId, 0)
                style = level + module + list_type + language
            para = Paragraph_(paraId, upParaId, style, no, content).__dict__
            paras.append(para)
            i += 1

    return paras,articleTittle,stuName,teacherName,majorName,gradTime


if __name__ == '__main__':
    paragraphs,articleTittle,stuName,teacherName,majorName,gradTime = read_word("demo3.docx")

    task = Task("e6c26921-3ec6-48b6-bb73-efd48cef969f",
                "428d81a2-30b7-4960-9535-c1c0e74e9677",
                "366775a1-f341-4e0b-ae45-382199d6c978",
                paragraphs,
                articleTittle,stuName,teacherName,majorName,gradTime)

    f = open("./test.json", 'w')
    f.write(json.dumps(task, default=task_2_json,
                       ensure_ascii=False,
                       sort_keys=True,
                       indent=4,
                       separators=(',', ': ')))
    f.close()
