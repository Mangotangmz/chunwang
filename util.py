import os
import zipfile
import os
import numpy as np
import pandas as pd

def set_back_color(s):
    #color = 'red' if s== None else 'red'
    color = 'yellow'
    return f"background-color:{color}"

def get_zip_file(input_path, result):
  """
  对目录进行深度优先遍历
  :param input_path:
  :param result:
  :return:
  """
  files = os.listdir(input_path)
  for file in files:
    if os.path.isdir(input_path + '/' + file):
      get_zip_file(input_path + '/' + file, result)
    else:
      result.append(input_path + '/' + file)


def zip_file_path(input_path, output_path, output_name):
  """
  压缩文件
  :param input_path: 压缩的文件夹路径
  :param output_path: 解压（输出）的路径
  :param output_name: 压缩包名称
  :return:
  """
  f = zipfile.ZipFile(output_path + '/' + output_name, 'w', zipfile.ZIP_DEFLATED)
  for root, dirnames, filenames in os.walk(input_path):
      file_path = root.replace(input_path, '')  # 去掉根路径，只对目标文件夹下的文件及文件夹进行压缩
      # 循环出一个个文件名
      for filename in filenames:
          f.write(os.path.join(root, filename), os.path.join(file_path, filename))
  return output_path + r"/" + output_name


def mkdir(path):
    # 引入模块

    # 去除首位空格
    path = path.strip()
    # 去除尾部 \ 符号
    path = path.rstrip("\\")
    # 判断路径是否存在
    # 存在     True
    # 不存在   False
    isExists = os.path.exists(path)
    # 判断结果
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
        print(path + ' 创建成功')
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        print(path + ' 目录已存在')
        return False


import os, sys


# 编写一个程序，能在某目录以及其所有子目录下查找文件名包含指定字符串的文件，并打印出相对路径。

def searchFile(key, startPath='.'):
    if not os.path.isdir(startPath):
        raise ValueError
    l = [os.path.join(startPath, x) for x in os.listdir(startPath)]  # 列出所有文件的绝对路径
    print(l)
    # listdir出来的相对路径 不能用于 isfile  abspath只能用在当前目录

    outlist = []
    for item in l :
        if key in str(item):
            outlist.append(item)


    return outlist



