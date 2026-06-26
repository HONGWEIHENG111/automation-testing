import threading

# 统一使用 Event 替代原有的布尔值，完美兼容单线程与多线程场景
# 这个对象将在所有 Runner 和监控循环中共享
STOP_EVENT = threading.Event()