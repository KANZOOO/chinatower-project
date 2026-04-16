import pickle
from collections import deque
import os
from core.config import settings
class yzm_queue:
    def __init__(self, max_length=10):
        self.max_length = max_length
        self.file_name = settings.resolve_path("spider\down\yzm_queue.pkl")
        self.queue = self._load_queue()

    def _load_queue(self):
        # 尝试从文件中加载队列
        if os.path.exists(self.file_name):
            with open(self.file_name, "rb") as file:
                return pickle.load(file)
        else:
            # 如果文件不存在，则初始化一个新的队列
            return deque(maxlen=self.max_length)

    def _save_queue(self):
        # 将队列保存到文件中
        with open(self.file_name, "wb") as file:
            pickle.dump(self.queue, file)

    def add_result(self, result):
        # 将结果添加到队列中
        self.queue.append(result)
        # 保存队列到文件
        self._save_queue()

    def get_queue(self):
        # 返回当前队列的内容
        return list(self.queue)

