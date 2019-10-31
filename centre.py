import threading
import time
import spider


def thread_function(name, version):
    spider.run(name, version)


if __name__ == "__main__":
    # lst = ["RFIB", "LSST", "RWFIB"]
    lst = ["RFIB"]

    version = "_All"
    for item in lst:
        x = threading.Thread(target=thread_function, args=(item, version,))
        x.start()
