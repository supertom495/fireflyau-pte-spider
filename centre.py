import threading
import time
import spider

def thread_function(name, version):
    spider.run(name, version)


if __name__ == "__main__":
    # lst = ["RFIB", "LSST", "RWFIB"]
    lst = ["RWFIB", "RFIB"]

    version = "5.2"
    for item in lst:
        x = threading.Thread(target=thread_function, args=(item, version,))
        x.start()
