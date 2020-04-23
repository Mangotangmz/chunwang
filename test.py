import threading

from main import recivecityFiles
from timer import timer

lock = threading.Lock()
threads = []
t1 = threading.Thread(target=recivecityFiles, args=(13,))
threads.append(t1)
# 循环检查是否有地市未反馈
t2 = threading.Thread(target=timer, args=('./files./operationlog/全省春旺地址核查处理日志(3.8-4.11).xlsx',))
threads.append(t2)
# 执行线程列表中的线程
for t in threads:
    # t.setDaemon(True)
    t.start()

for p in threads:
    p.join()
print('3')