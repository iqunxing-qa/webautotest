#coding=utf-8
#############################
#获取字符串中第N次出现的位置#
#############################
def findStr(string, subStr, findCnt):
    listStr = string.split(subStr,findCnt)
    print listStr
    if len(listStr) <= findCnt:    #分割完后的字符串的长度（分割段）与要求出现的次数比较
        return -1
    return len(string)-len(listStr[-1])-len(subStr)#len(listStr[-1])最后的一个集合里面字符串的长度  ，len(subStr)  减去本身的长度
#################第二种方法######################
def findSubStr(substr, str, i):
    count = 0
    while i > 0:                   #循环来查找
        index = str.find(substr)
        if index == -1:
            return -1
        else:
            str = str[index+1:]   #第一次出现该字符串后后面的字符
            i -= 1
            count = count + index + 1   #位置数总加起来
    return count - 1