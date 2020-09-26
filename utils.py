class Paragraph_:
    def __init__(self, paraId, upParaId, style, no, content):
        self.paraId = paraId
        self.upParaId = upParaId
        self.style = style
        self.no = no
        self.content = content


class element_:
    def __init__(self, paraId, style, no):
        self.paraId = paraId
        self.style = style
        self.no = no


class Stack(object):

    def __init__(self):
        # 创建空列表实现栈
        self.__list = []

    def is_empty(self):
        # 判断是否为空
        return self.__list == []

    def push(self, item):
        # 压栈，添加元素
        self.__list.append(item)

    def pop(self):
        # 弹栈，弹出最后压入栈的元素
        if self.is_empty():
            return
        else:
            return self.__list.pop()

    def top(self):
        # 取最后压入栈的元素
        if self.is_empty():
            return
        else:
            return self.__list[-1]


class Task:
    def __init__(self, userId, articleId, ruleSheetId,paragraphs, articleTittle, stuName, teacherName, majorName, gradTime):
        self.userId = userId
        self.articleId = articleId
        self.ruleSheetId = ruleSheetId
        self.paragraphs = paragraphs
        self.articleTittle = articleTittle
        self.stuName = stuName
        self.teacherName = teacherName
        self.majorName = majorName
        self.gradTime = gradTime


def task_2_json(obj):
    return {
        "userId": obj.userId,
        "articleId": obj.articleId,
        "ruleSheetId": obj.ruleSheetId,
        "articleTittle": obj.articleTittle,
        "stuName": obj.stuName,
        "teacherName": obj.teacherName,
        "majorName": obj.majorName,
        "gradTime": obj.gradTime,
        "paragraphs": obj.paragraphs
    }
