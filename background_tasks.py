"""Run data work off-thread, deliver only the latest result on the UI thread."""

from concurrent.futures import ThreadPoolExecutor
from queue import Queue, Empty
import threading

from trace_store import Cancelled


class BackgroundTasks:
    def __init__(self):
        self.executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix='trace-query')
        self.results = Queue()
        self.latest = {}
        self.closed = False

    def submit(self, channel, work, done, error):
        previous = self.latest.get(channel)
        if previous:
            previous.set()
        token = threading.Event()
        self.latest[channel] = token

        def run():
            try:
                if token.is_set():
                    return
                result = work(token)
                self.results.put((channel, token, done, result))
            except Cancelled:
                pass
            except Exception as exception:
                self.results.put((channel, token, error, exception))
        self.executor.submit(run)
        return token

    def drain(self):
        while True:
            try:
                channel, token, callback, value = self.results.get_nowait()
            except Empty:
                break
            if not self.closed and self.latest.get(channel) is token and not token.is_set():
                callback(value)

    def cancel(self, channel):
        token = self.latest.get(channel)
        if token:
            token.set()

    def close(self):
        self.closed = True
        for token in self.latest.values():
            token.set()
        self.executor.shutdown(wait=False, cancel_futures=True)
